"""First-passport discovery from project's *existing* historical sources.

Mocks only the already separately-tested first-appendix native layout guard.
Every fixture is synthetic, with no customer data or production registry.
"""
import unittest
from datetime import date
from pathlib import Path
from tempfile import TemporaryDirectory
from types import SimpleNamespace
from unittest.mock import patch
from zipfile import ZipFile

from openpyxl import Workbook

from production_app.first_passport import (
    discover_first_passport,validate_source_pins,
)
from production_app.storage import Stop


def project():
    return SimpleNamespace(
        name='Гараж Казахстан',first_number='1/2026',
        first_date=date(2026,9,9),object_name='СИНТЕТИКА'
    )


def docx(path,passport_no='1/2026'):
    xml=(
      '<?xml version="1.0" encoding="UTF-8"?>'
      '<w:document xmlns:w="http://schemas.openxmlformats.org/'
      'wordprocessingml/2006/main"><w:body><w:p><w:r>'
      '<w:t>Акт № '+passport_no+'</w:t>'
      '</w:r></w:p></w:body></w:document>'
    )
    with ZipFile(path,'w') as z:
        z.writestr('word/document.xml',xml)


def xlsx(path,number='1/2026'):
    w=Workbook()
    w.active['E4']=number+' от '
    w.save(path)
    w.close()


class FirstPassportSources(unittest.TestCase):
    def setUp(self):
        self.tmp=TemporaryDirectory()
        self.root=Path(self.tmp.name)
        self.other=(self.root/'ТЭМПО 1'/'Паспорт')
        self.other.mkdir(parents=True)
        self.passport=(self.root/'Гараж казахстан'/'Паспорт')
        self.passport.mkdir(parents=True)
        self.app=self.passport/'1_2026 - Приложение № 1.xlsx'
        xlsx(self.app)
        for name in ('ВИК','УЗК','Документ о качестве'):
            docx(self.passport/('1_2026 - '+name+'.docx'))
        self.akz=self.passport/'Акт № 4-2026.xlsx'
        xlsx(self.akz)
        self.akz_appendix=self.passport/'1_2026 - Приложение к АКЗ.xlsx'
        xlsx(self.akz_appendix)
        (self.passport.parent/'Справочник марок.xlsx').write_bytes(b'FAKE_CATALOG')

    def tearDown(self):
        self.tmp.cleanup()

    def scan(self,**kwargs):
        # Already-tested native first-appendix validator must independently
        # verify first number/date/object; this mock isolates source discovery.
        with patch('production_app.first_passport.inspect_template',
                   return_value=object()):
            return discover_first_passport(self.root,project(),**kwargs)

    def test_detects_exact_first_sources_in_original_passport_folder(self):
        result=self.scan(require_all=True)
        self.assertEqual(set(result['sources']),
            {'appendix','vik','uzk','quality','akz_appendix0','akz0'})
        self.assertEqual(Path(result['sources']['appendix']['path']),self.app)
        self.assertEqual(Path(result['sources']['akz0']['path']),self.akz)
        self.assertEqual(result['first_number'],'1/2026')
        self.assertEqual(result['status'],'READ_ONLY_FIRST_PASSPORT')
        self.assertTrue(validate_source_pins(result))

    def test_no_mixing_with_other_projects_or_later_passports(self):
        docx(self.other/'1_2026 - УЗК.docx')
        folder=self.passport/'10 от 07.10.2026'
        folder.mkdir()
        docx(folder/'10_2026 - ВИК.docx',passport_no='10/2026')
        result=self.scan()
        self.assertEqual(Path(result['sources']['uzk']['path']).parent,
                         self.passport)

    def test_casefold_project_folder_unique_or_stop(self):
        (self.root/'ГАРАЖ КАЗАХСТАН').mkdir()
        with self.assertRaisesRegex(Stop,'единственная папка'):
            self.scan()

    def test_missing_first_word_file_is_reported_without_fallback(self):
        (self.passport/'1_2026 - ВИК.docx').unlink()
        result=self.scan()
        self.assertIn('vik',result['missing'])
        self.assertNotIn('vik',result['sources'])
        with self.assertRaisesRegex(Stop,'ВИК|vik'):
            self.scan(require_all=True)

    def test_missing_first_appendix_refuses_lookalike_second(self):
        self.app.unlink()
        xlsx(self.passport/'2_2026 - Приложение № 1.xlsx',
             number='2/2026')
        result=self.scan()
        self.assertIn('appendix',result['missing'])

    def test_duplicate_first_source_stops_instead_of_guessing(self):
        second=self.passport/'1 от 09.09.2026'
        second.mkdir()
        docx(second/'1_2026 - ВИК.docx')
        with self.assertRaisesRegex(Stop,'Несколько документов'):
            self.scan()

    def test_stale_word_act_with_number_2_refused(self):
        docx(self.passport/'1_2026 - ВИК.docx',passport_no='2/2026')
        with self.assertRaisesRegex(Stop,'другой номер'):
            self.scan()

    def test_altered_issued_document_fails_source_pin(self):
        result=self.scan()
        with self.akz.open('ab') as f:
            f.write(b'changed')
        with self.assertRaisesRegex(Stop,'изменился'):
            validate_source_pins(result)

    def test_akz_4_filename_can_be_first_only_when_inner_number_matches(self):
        xlsx(self.akz,number='4/2026')
        with self.assertRaisesRegex(Stop,'другой номер'):
            self.scan()

    def test_catalog_is_only_an_optional_candidate(self):
        (self.passport.parent/'Справочник марок.xlsx').unlink()
        result=self.scan()
        self.assertIsNone(result['catalog'])
        self.assertTrue(validate_source_pins(result))

    def test_does_not_create_or_modify_any_first_source(self):
        before={p:p.read_bytes() for p in self.passport.iterdir() if p.is_file()}
        self.scan()
        self.assertEqual({p:p.read_bytes() for p in before},before)


if __name__=='__main__':
    unittest.main(verbosity=2)
