import tempfile
import tkinter as tk
import unittest
from pathlib import Path
from types import SimpleNamespace

from openpyxl import Workbook

from production_app.automatic_acts import (
    PROFILE_HEADINGS, calculate_design_characteristics, profile_standards,
)
from production_app.geometry_selection import GeometrySelection
from production_app.gui import App
from production_app.storage import Stop
from profiles_stage4.model import Catalog, normalize_profile


class ChannelRecognition(unittest.TestCase):
    def test_prefix_preserves_series_and_rejects_non_channel_alias(self):
        self.assertEqual(normalize_profile('швеллер\u00a016 п'), '16П')
        self.assertEqual(normalize_profile('Швеллер 16У'), '16У')
        self.assertEqual(normalize_profile('Швеллер 16Э'), '16Э')
        records = [
            {'record_id': key, 'display': key, 'key': key, 'standard': standard,
             'family': family, 'geometry': {'h': 160, 'b': 64, 's': 5, 't': 8.4},
             'source': 'FIXTURE'}
            for key, standard, family in (
                ('16П', 'ГОСТ 8240-97', 'Швеллер'),
                ('16У', 'ГОСТ 8240-97', 'Швеллер'),
                ('20Б1', 'ГОСТ 26020-83', 'Двутавр'),
            )
        ]
        catalog = Catalog(records)
        self.assertEqual(catalog.lookup('Швеллер 16У')['profile']['key'], '16У')
        self.assertEqual(catalog.lookup('Швеллер 20Б1')['status'], 'NOT_FOUND')
        self.assertEqual(catalog.lookup('Швеллер 16Н')['status'], 'NOT_FOUND')

    def test_selection_and_design_use_the_approved_channel_rule(self):
        with tempfile.TemporaryDirectory() as temporary:
            path = Path(temporary) / 'Профили.xlsx'
            book = Workbook()
            sheet = book.active
            sheet.title = 'Профили'
            sheet.append(PROFILE_HEADINGS)
            sheet.append(('16П', '16П', 'Швеллер', 'ГОСТ 8240-97', 160, 64, None, 5, 8.4,
                          None, None, 'FIXTURE', None, 'УТВЕРЖДЕНО', 'ШВЕЛЛЕР_16П_1', None))
            sheet.append(('20Б1', '20Б1', 'Двутавр', 'ГОСТ 26020-83', 200, 100, None, 5, 8,
                          None, None, 'FIXTURE', None, 'УТВЕРЖДЕНО', 'ДВУТАВР_ПРОКАТНЫЙ_1', None))
            book.save(path)
            book.close()

            selection = GeometrySelection(path, {}, 'DEV')
            self.assertEqual(profile_standards(path, 'Швеллер 16П'), ['ГОСТ 8240-97'])
            self.assertEqual(profile_standards(path, 'Швеллер 20Б1'), [])
            self.assertEqual(selection.options('Швеллер 20Б1'), [])
            self.assertEqual(selection.resolve('M1', 'Швеллер 16П'), 'ГОСТ 8240-97')
            vom = SimpleNamespace(rows=(('M1',),))
            facts = {'marks': [{'mark': 'M1', 'profile': 'Швеллер 16П',
                                'standard': 'ГОСТ 8240-97', 'uzk_quantity': 2}]}
            rows = calculate_design_characteristics(vom, facts, path)
            self.assertEqual([(r['length_mm'], r['thickness_mm']) for r in rows],
                             [('320', '5'), ('256', '8.4')])
            self.assertTrue(all(r['profile'] == 'Швеллер 16П' for r in rows))
            facts['marks'][0].update(profile='Швеллер 20Б1', standard='ГОСТ 26020-83')
            with self.assertRaises(Stop):
                calculate_design_characteristics(vom, facts, path)


class FactsErrorDialog(unittest.TestCase):
    def test_long_message_closes_and_returns_grab_to_facts(self):
        try:
            root = tk.Tk()
        except tk.TclError as exc:
            self.skipTest(str(exc))
        self.addCleanup(root.destroy)
        app = App.__new__(App)
        log = []
        app.write = log.append
        facts = tk.Toplevel(root)
        facts.grab_set()
        text = '\n'.join(f'Ошибка {index}' for index in range(80))

        def descendants(widget):
            for child in widget.winfo_children():
                yield child
                yield from descendants(child)

        for use_close_button in (True, False):
            failures = []

            def inspect_and_close():
                dialog = next((w for w in facts.winfo_children() if isinstance(w, tk.Toplevel)), None)
                try:
                    self.assertIsNotNone(dialog)
                    self.assertEqual(facts.grab_current(), dialog)
                    self.assertEqual(str(dialog.transient()), str(facts))
                    editor = next(w for w in descendants(dialog) if isinstance(w, tk.Text))
                    self.assertEqual(editor.get('1.0', 'end-1c'), text)
                    self.assertLess(editor.yview()[1], 1.0)
                    editor.yview_moveto(1.0)
                    self.assertEqual(editor.yview()[1], 1.0)
                    button = next(w for w in descendants(dialog) if w.winfo_class() == 'TButton')
                    self.assertTrue(button.winfo_viewable())
                    if use_close_button:
                        button.invoke()
                    else:
                        dialog.tk.call(dialog.protocol('WM_DELETE_WINDOW'))
                    self.assertFalse(dialog.winfo_exists())
                except Exception as exc:
                    failures.append(exc)
                finally:
                    if dialog is not None and dialog.winfo_exists():
                        dialog.destroy()

            root.after(100, inspect_and_close)
            app.error(text, parent=facts)
            if failures:
                raise failures[0]
            self.assertEqual(facts.grab_current(), facts)
        self.assertEqual(log, [text, text])
