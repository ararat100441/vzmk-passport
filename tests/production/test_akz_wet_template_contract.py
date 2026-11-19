"""Bounded validation of already exported wet tokens; no approved form fixtures."""
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from openpyxl import Workbook

from production_app import templates
from production_app.storage import Stop
from production_app.working_facts import selected_form


class WetTemplateContract(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.path = Path(self.temp.name) / 'synthetic-contract.xlsx'
        self.fields = {
            'NUMBER', 'CONTRACT', 'CUSTOMER', 'OBJECT', 'CIPHER',
            'AKZ_DATE', 'AIR', 'SURFACE', 'HUMIDITY', 'DEW',
        }
        self.fields.update(f'LAYER_{n}_{key}' for n in (1, 2, 3)
                           for key in ('MIN', 'MAX', 'DRY_AVG'))
        self.fields.update(f'AKZ_LAYER_{n}_{key}' for n in (1, 2, 3)
                           for key in ('MATERIAL', 'PAINT_BATCH', 'HARDENER_BATCH',
                                       'COLOR', 'APPLICATION_DATE', 'TIME_START',
                                       'TIME_END', 'WET_AVG'))

    def validate_fields(self, extra=(), remove=()):
        book = Workbook()
        for row, key in enumerate(sorted((self.fields | set(extra)) - set(remove)), 1):
            book.active.cell(row, 1, '{{' + key + '}}')
        book.save(self.path)
        book.close()
        appendix = self.path.with_name('synthetic-appendix.xlsx')
        book = Workbook()
        for row, value in enumerate(('{{NUMBER}} {{OBJECT}} {{AKZ_DATE}}',
                                      '{{AKZ_MARKS}}', 'Итого'), 1):
            book.active.cell(row, 1, value)
        book.save(appendix)
        book.close()
        profile = {'templates': {role: {'path': str(self.path)} for role in
                                 ('appendix', 'vik', 'uzk', 'quality', 'akz0', 'akz_appendix0')}}
        profile['templates']['akz_appendix0'] = {'path': str(appendix)}
        for role in templates.WORD_REQUIRED:
            profile['templates'][role] = {'path': str(self.path.with_suffix('.docx'))}
        # These unrelated documents are outside this Excel-contract unit test.
        word = set().union(*templates.WORD_REQUIRED.values()) | {'UZK_TABLE'}
        with patch.object(templates, 'inspect_template'), \
                patch.object(templates, 'inspect_word_fields', return_value=(word, set())):
            self.assertTrue(templates.validate(profile, object(), [0], training=True))

    def test_exported_wet_tokens_are_accepted(self):
        wet = {f'LAYER_{n}_{key}' for n in (1, 2, 3)
               for key in ('WET_AVG', 'WET_TOTAL')}
        self.validate_fields(wet | {'WET_TOTAL'})

    def test_unknown_wet_layer_still_rejected(self):
        with self.assertRaisesRegex(Stop, 'Неподдержанное поле'):
            self.validate_fields({'LAYER_4_WET_AVG'})

    def test_wet_tokens_do_not_replace_cumulative_dry(self):
        with self.assertRaisesRegex(Stop, 'среднюю сухую'):
            self.validate_fields({'LAYER_3_WET_AVG'}, {'LAYER_3_DRY_AVG'})

    def test_wet_tokens_do_not_replace_confirmed_layer_requisites(self):
        with self.assertRaisesRegex(Stop, 'все поля ЛКМ'):
            self.validate_fields({'LAYER_1_WET_AVG'}, {'AKZ_LAYER_2_PAINT_BATCH'})

    def test_first_passport_inventory_still_prevents_clean_fallback(self):
        profile = {'templates': {'akz0': {'path': str(self.path)}},
                   'first_passport_sources': {'first_number': '1/2026', 'sources': {}}}
        with self.assertRaisesRegex(Stop, 'Другая форма автоматически не подставлена'):
            selected_form(profile, 'akz0')


if __name__ == '__main__':
    unittest.main()
