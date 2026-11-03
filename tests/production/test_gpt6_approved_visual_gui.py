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
        self.local = patch('production_app.gui.local_root',
                           return_value=Path(self.temp.name))
        self.local.start()
        self.addCleanup(self.local.stop)
        self.app = gui.App(self.root)
        self.root.update_idletasks()

    def widgets(self, parent=None):
        parent = self.root if parent is None else parent
        for child in parent.winfo_children():
            yield child
            yield from self.widgets(child)

    def test_existing_commands_and_selection_widgets_are_preserved(self):
        self.assertEqual(len(self.app.buttons), 1)
        self.assertEqual(self.app.buttons[0].cget('text'), 'Сформировать пакет')
        self.assertEqual(self.app.buttons[0].cget('style'), 'Accent.TButton')
        self.assertEqual(self.app.projects.cget('state'), 'readonly')
        self.assertEqual(self.app.numbers.cget('state'), 'readonly')
        self.assertEqual(self.app.source_numbers.cget('state'), 'readonly')
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
