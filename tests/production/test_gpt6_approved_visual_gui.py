"""Visual-only GUI acceptance for the approved VZMK desktop design.

These checks require an interactive Windows desktop/Tk installation, but do
not open Office, read production registries, or issue any real documents.
"""
import os
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import tkinter as tk
from tkinter import ttk

from production_app import gui


@unittest.skipUnless(os.name == 'nt', 'Native Windows Tk GUI verification')
class ApprovedVisualLayout(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = tk.Tk()
        self.root.withdraw()
        self.addCleanup(self.root.destroy)
        self.scheduled=[]
        self.local = patch('production_app.gui.local_root',
                           return_value=Path(self.temp.name))
        self.local.start()
        self.addCleanup(self.local.stop)
        # Never display a modal error over a hidden test root: fail promptly.
        # Do not run delayed settings dialogs under a withdrawn test root.
        # Check that they were scheduled, without leaving Tk callbacks alive.
        def capture_after(delay, callback, *args):
            self.scheduled.append((delay, callback, args))
            return 'captured-schedule'
        with patch.object(self.root, 'after', side_effect=capture_after), \
             patch('production_app.gui.messagebox.showerror',
                   side_effect=AssertionError('Unexpected GUI startup error')):
            self.app = gui.App(self.root)
        self.root.update_idletasks()

    def widgets(self, parent=None):
        parent = self.root if parent is None else parent
        for child in parent.winfo_children():
            yield child
            yield from self.widgets(child)

    def test_first_run_without_registry_does_not_select_project_or_show_error(self):
        # This is the CODEX51 regression: empty ttk.Combobox values can be ''
        # and Python evaluates '' in '' as True unless the name is checked.
        self.assertFalse(self.app.cfg.get('registry'))
        self.assertEqual(self.app.project.get(), '')
        self.assertFalse(self.app.projects['values'])
        self.assertIsNone(self.app.plan)
        self.assertIn('Выберите рабочие пути',
                      self.app.progress.cget('text'))
        self.assertEqual(len(self.scheduled), 1)
        self.assertEqual(self.scheduled[0][0], 300)
        self.assertEqual(self.scheduled[0][1].__name__, 'settings')

    def test_configured_project_still_uses_existing_selection_callback(self):
        # The first-run fix must not suppress selection for a valid profile.
        from production_app.storage import save
        folder = Path(self.temp.name)
        save(folder / 'rc3-marked-settings.json', {
            'registry': 'CONFIGURED_REGISTRY.xlsx',
            'profiles': {'TEST': {}},
            'last_project': 'TEST',
        })
        def fake_refresh(app):
            app.projects['values'] = ('TEST',)
        configured_root = tk.Toplevel(self.root)
        configured_root.withdraw()
        self.addCleanup(configured_root.destroy)
        with patch.object(gui.App, 'refresh', new=fake_refresh), \
             patch.object(gui.App, 'select_project', autospec=True) as selected, \
             patch('production_app.gui.messagebox.showerror',
                   side_effect=AssertionError('Unexpected configured startup error')):
            configured = gui.App(configured_root)
        self.assertEqual(configured.project.get(), 'TEST')
        selected.assert_called_once_with(configured)

    def test_existing_commands_and_selection_widgets_are_preserved(self):
        self.assertEqual(len(self.app.buttons), 1)
        self.assertEqual(self.app.buttons[0].cget('text'), 'Сформировать пакет')
        self.assertEqual(self.app.buttons[0].cget('style'), 'Accent.TButton')
        # Tcl may return a Tcl_Obj instead of str for the index state.
        self.assertEqual(str(self.app.projects.cget('state')), 'readonly')
        self.assertEqual(str(self.app.numbers.cget('state')), 'readonly')
        self.assertEqual(str(self.app.source_numbers.cget('state')), 'readonly')
        self.assertEqual(self.app.new_widget.cget('text'), 'Новый паспорт')
        self.assertIsNotNone(self.app.shipment_entry)
        self.assertIsNotNone(self.app.log)
        actions = {button.cget('text') for button in self.widgets()
                   if isinstance(button, ttk.Button)}
        self.assertEqual(actions, {
            'Настройки путей', 'Обновить реестр', 'Отгрузочная ВОМ',
            'Фактические результаты', 'Открыть документы',
            'Сформировать пакет',
        })

    def test_three_sections_and_one_operation_log(self):
        titles = {w.cget('text') for w in self.widgets()
                  if isinstance(w, ttk.Labelframe)}
        self.assertEqual(titles, {
            '01   Данные паспорта',
            '02   Отгрузочная ведомость',
            '03   Документы и контроль',
            'Журнал операций',
        })

    def test_visual_summary_preserves_source_values_and_unknown_area(self):
        self.app.vom_summary.set(
            '3 строк / 2 уникальных марок / 3 изделий / 842.92 кг / 26.877 м²'
        )
        self.assertEqual(
            [self.app.vom_kpis[k].get() for k in
             ('marks', 'pieces', 'mass', 'area')],
            ['2', '3', '842.92 кг', '26.877 м²'],
        )
        self.app.vom_summary.set(
            '3 строк / 2 уникальных марок / 3 изделий / 842.92 кг / площадь не указана'
        )
        self.assertEqual(self.app.vom_kpis['area'].get(), 'не указана')
        self.app.vom_summary.set('Строки ВОМ ещё не введены')
        self.assertTrue(all(v.get() == '—' for v in self.app.vom_kpis.values()))

    def test_visual_only_theme_and_log_scrollbar(self):
        styles = ttk.Style(self.root)
        self.assertEqual(
            styles.configure('Accent.TButton', 'background'),
            gui.GUI_PRIMARY,
        )
        self.assertEqual(self.root.cget('background'), gui.GUI_BACKGROUND)
        self.assertTrue(any(isinstance(w, ttk.Scrollbar)
                            for w in self.widgets()))
        self.app.write('ТЕСТ: только визуальное оформление')
        self.assertIn('ТЕСТ: только визуальное оформление',
                      self.app.log.get('1.0', 'end'))


if __name__ == '__main__':
    unittest.main()
