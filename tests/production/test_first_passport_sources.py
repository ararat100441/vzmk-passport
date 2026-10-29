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
    discover_first_passport,discover_first_passport_for_config,
    validate_source_pins,
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
        # Windows cannot create two sibling names differing only by case.
        # Internal whitespace gives a real second path with the same _key.
        (self.root/'ГАРАЖ  КАЗАХСТАН').mkdir()
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

    def test_deployed_rc3_templates_and_real_document_roots_are_distinct(self):
        # The deployed screenshot has 'documents'=.../Templates,
        # while actual FIRST passports reside at cfg['output'].
        cleaned=self.root/'CleanTemplates'
        (cleaned/'Гараж Казахстан').mkdir(parents=True)
        cfg={'documents':str(cleaned),'output':str(self.root)}
        with patch('production_app.first_passport.inspect_template',
                   return_value=object()):
            result=discover_first_passport_for_config(
                cfg,project(),require_all=True)
        self.assertEqual(Path(result['passport_root']),self.passport)
        self.assertEqual(Path(result['sources']['appendix']['path']),self.app)

    def test_explicit_archive_root_overrides_unrelated_output(self):
        cleaned=self.root/'CleanTemplates'
        cleaned.mkdir()
        unrelated=self.root/'EmptyOutput'
        unrelated.mkdir()
        cfg={'documents':str(cleaned),'output':str(unrelated),
             'first_passport_documents':str(self.root)}
        with patch('production_app.first_passport.inspect_template',
                   return_value=object()):
            result=discover_first_passport_for_config(cfg,project())
        self.assertEqual(Path(result['passport_root']),self.passport)

    def test_missing_archive_in_selected_paths_refuses_random_fallback(self):
        cleaned=self.root/'CleanTemplates'
        (cleaned/'Гараж Казахстан').mkdir(parents=True)
        unrelated=self.root/'EmptyOutput'
        unrelated.mkdir()
        cfg={'documents':str(cleaned),'output':str(unrelated)}
        with self.assertRaisesRegex(Stop,'Архив первых паспортов'):
            discover_first_passport_for_config(cfg,project())

    def test_two_potential_archives_are_ambiguous_until_user_selects_one(self):
        cleaned=self.root/'CleanTemplates'
        archive=cleaned/'Гараж Казахстан'/'Паспорт'
        archive.mkdir(parents=True)
        # Two actual archives selected by user: must not silently choose.
        cfg={'documents':str(cleaned),'output':str(self.root)}
        with self.assertRaisesRegex(Stop,'корневых папок: 2'):
            discover_first_passport_for_config(cfg,project())

    def test_explicit_archive_rejects_missing_project_even_when_output_has_one(self):
        cleaned=self.root/'CleanTemplates'
        cleaned.mkdir()
        cfg={'documents':str(cleaned),'output':str(self.root),
             'first_passport_documents':str(cleaned)}
        with self.assertRaisesRegex(Stop,'единственная папка'):
            discover_first_passport_for_config(cfg,project())

    def test_profile_uses_first_appendix_not_cached_surrogate(self):
        from production_app import model
        prepared=self.passport.parent/'Prepared'
        prepared.mkdir()
        templates={}
        for role,filename in (
            ('appendix','old_cached_app.xlsx'),('vik','clean_vik.docx'),
            ('uzk','clean_uzk.docx'),('quality','clean_quality.docx'),
        ):
            path=prepared/filename
            path.write_bytes(role.encode('utf8'))
            from production_app.storage import sha
            templates[role]={'path':str(path),'sha256':sha(path)}
        cfg={
            'documents':str(self.root),
            'registry':str(self.root/'sample-registry.xlsx'),
            'auto_first_passport':True,
            'profiles':{
                project().name:{
                    'root':str(self.passport.parent),
                    'templates':templates
                }
            }
        }
        with patch('production_app.first_passport.inspect_template',
                   return_value=object()):
            with patch('production_app.model.read_registry',return_value=object()):
                with patch('production_app.model.select',return_value=project()):
                    actual=model.profile(cfg,project().name)
        self.assertEqual(Path(actual['templates']['appendix']['path']),self.app)
        for role in ('vik','uzk','quality'):
            self.assertEqual(actual['templates'][role]['path'],templates[role]['path'])
            # Issued old acts never silently become approved empty templates.
            self.assertNotEqual(actual['templates'][role]['path'],
                                actual['first_passport_sources']['sources'][role]['path'])
        self.assertEqual(cfg['profiles'][project().name]['templates'],templates)

    def test_auto_mode_does_not_require_old_appendix_assignment(self):
        from production_app import model
        from production_app.storage import sha
        prepared=self.passport.parent/'Prepared'
        prepared.mkdir()
        templates={}
        for role in ('vik','uzk','quality'):
            path=prepared/(role+'.docx')
            path.write_bytes(role.encode())
            templates[role]={'path':str(path),'sha256':sha(path)}
        cfg={'documents':str(self.root),'registry':str(self.root/'registry.xlsx'),
             'auto_first_passport':True,
             'profiles':{project().name:{
                 'root':str(self.passport.parent),'templates':templates,
             }}}
        with patch('production_app.first_passport.inspect_template',
                   return_value=object()):
            with patch('production_app.model.read_registry',return_value=object()):
                with patch('production_app.model.select',return_value=project()):
                    result=model.profile(cfg,project().name)
        self.assertEqual(Path(result['templates']['appendix']['path']),self.app)
        self.assertEqual(cfg['profiles'][project().name]['templates'],templates)

    def test_auto_mode_ignores_stale_old_cached_appendix(self):
        from production_app import model
        from production_app.storage import sha
        prepared=self.passport.parent/'Prepared'
        prepared.mkdir()
        templates={}
        for role in ('appendix','vik','uzk','quality'):
            path=prepared/(role+'.xlsx')
            path.write_bytes(role.encode())
            templates[role]={'path':str(path),'sha256':sha(path)}
        templates['appendix']['sha256']='outdated-first-appendix-checksum'
        cfg={'documents':str(self.root),'registry':str(self.root/'registry.xlsx'),
             'auto_first_passport':True,
             'profiles':{project().name:{
                 'root':str(self.passport.parent),'templates':templates,
             }}}
        with patch('production_app.first_passport.inspect_template',
                   return_value=object()):
            with patch('production_app.model.read_registry',return_value=object()):
                with patch('production_app.model.select',return_value=project()):
                    result=model.profile(cfg,project().name)
        self.assertEqual(Path(result['templates']['appendix']['path']),self.app)
        self.assertNotEqual(result['templates']['appendix']['sha256'],
                            'outdated-first-appendix-checksum')

    def test_explicit_manual_mode_keeps_existing_approved_appendix(self):
        from production_app import model
        prepared=self.passport.parent/'Prepared'
        prepared.mkdir()
        templates={}
        from production_app.storage import sha
        for role in ('appendix','vik','uzk','quality'):
            p=prepared/(role+'.xlsx')
            p.write_bytes(role.encode())
            templates[role]={'path':str(p),'sha256':sha(p)}
        cfg={
            'documents':str(self.root),'auto_first_passport':False,
            'profiles':{project().name:{
                'root':str(self.passport.parent),'templates':templates
            }}
        }
        actual=model.profile(cfg,project().name)
        self.assertEqual(actual['templates']['appendix'],templates['appendix'])
        self.assertNotIn('first_passport_sources',actual)

    def test_does_not_create_or_modify_any_first_source(self):
        before={p:p.read_bytes() for p in self.passport.iterdir() if p.is_file()}
        self.scan()
        self.assertEqual({p:p.read_bytes() for p in before},before)


if __name__=='__main__':
    unittest.main(verbosity=2)
