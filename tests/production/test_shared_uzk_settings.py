"""Contracts: common read-only UZK registry, three visible paths, dated first."""

import tempfile
import unittest
from pathlib import Path
from datetime import date
from types import SimpleNamespace
from unittest.mock import patch
from contextlib import ExitStack

from openpyxl import Workbook,load_workbook

from production_app import core,gui
from production_app.storage import Stop,sha
from production_app.shared_uzk import (
    read_shared_uzk,planned_control,verify_planned_quantities,find_shared_uzk
)
from production_app.path_config import apply_settings,visible_settings
from production_app.prepared_forms import prepared_project
from production_app.model import output_passport_directory
from production_app.first_passport import _first_roots


def workbook(path,rows):
    w=Workbook()
    w.active.title='УЗК'
    for row in rows:w.active.append(row)
    w.save(path);w.close()


class SharedUzk(unittest.TestCase):
    def setUp(self):
        self.tmp=tempfile.TemporaryDirectory()
        self.root=Path(self.tmp.name)
        self.file=self.root/'УЗК.xlsm'
        self.rows=[
            ['Заказ','Марка элемента','Кол-во (шт)'],
            ['Гараж Казахстан','Бп2-5',None],
            ['Гараж Казахстан','Бп2-18',2],
            ['Гараж Казахстан','Бп2-18',2],
            ['ТЭМПО 1','Бп2-5',5],
            ['ТЭМПО 1','Т1',3],
        ]
        workbook(self.file,self.rows)
    def tearDown(self):self.tmp.cleanup()

    def vom(self):
        return SimpleNamespace(rows=(
            ('Бп2-5','Балка',3,1,3,1,3),
            ('Бп2-18','Балка',1,1,1,1,1),
            ('Бп2-18','Балка',5,1,5,1,5),
            ('Не из реестра','Балка',1,1,1,1,1),
        ))

    def test_blank_quantity_is_one_and_duplicates_do_not_double(self):
        start=sha(self.file)
        values=read_shared_uzk(self.file)
        self.assertEqual(values[('гараж казахстан','бп2-5')]['remaining'],1)
        self.assertEqual(values[('гараж казахстан','бп2-18')]['remaining'],2)
        self.assertEqual(values[('гараж казахстан','бп2-18')]['rows'],[3,4])
        self.assertEqual(sha(self.file),start)

    def test_selected_project_matches_exact_marks_and_capped_quantity(self):
        expected=planned_control(self.file,'Гараж Казахстан',self.vom())
        self.assertEqual(expected['Бп2-5']['uzk_quantity'],1)
        self.assertEqual(expected['Бп2-18']['uzk_quantity'],2)
        self.assertEqual(expected['Бп2-18']['vom_quantity'],6)
        self.assertEqual(expected['Не из реестра']['uzk_quantity'],0)
        self.assertFalse(expected['Не из реестра']['matched'])
        self.assertEqual(len(expected),3)
        self.assertNotIn('result',expected['Бп2-18'])

    def test_same_mark_of_another_project_never_counted(self):
        result=planned_control(self.file,'ТЭМПО 1',self.vom())
        self.assertEqual(result['Бп2-5']['uzk_quantity'],3)
        self.assertEqual(result['Бп2-18']['uzk_quantity'],0)

    def test_conflicting_duplicate_is_refused_without_touching_source(self):
        self.rows[3]=['Гараж Казахстан','Бп2-18',9]
        workbook(self.file,self.rows)
        before=sha(self.file)
        with self.assertRaisesRegex(Stop,'противоречивые остатки'):
            read_shared_uzk(self.file)
        self.assertEqual(sha(self.file),before)

    def test_false_zero_invalid_and_formula_invalid(self):
        for value in (0,-1,1.25,True,'=1+1'):
            with self.subTest(value=value):
                self.rows[1]=['Гараж Казахстан','Бп2-5',value]
                workbook(self.file,self.rows)
                with self.assertRaises(Stop):read_shared_uzk(self.file)

    def test_wrong_headers_and_missing_sheet_stop(self):
        self.rows[0]=['Заказ','Марка','Количество']
        workbook(self.file,self.rows)
        with self.assertRaisesRegex(Stop,'заголовки'):
            read_shared_uzk(self.file)

    def test_planned_quantities_cannot_be_changed_to_avoid_control(self):
        expected=planned_control(self.file,'Гараж Казахстан',self.vom())
        facts={'marks':[
            {'mark':mark,'uzk_quantity':row['uzk_quantity']}
            for mark,row in expected.items()
        ]}
        self.assertTrue(verify_planned_quantities(facts,expected))
        facts['marks'][0]['uzk_quantity']=0
        with self.assertRaisesRegex(Stop,'должно быть'):
            verify_planned_quantities(facts,expected)

    def test_third_visible_setting_shared_for_all_projects(self):
        source=self.root/'ВЗМК (2).xlsx'
        source.write_bytes(b'NO REAL REGISTRY')
        docs=self.root/'Документы';docs.mkdir()
        forms=self.root/'Templates';forms.mkdir()
        cfg={'documents':str(forms),'registry':str(source),
             'output':str(docs),'profiles':{'Гараж Казахстан':{'root':str(forms)}}}
        shown=visible_settings(cfg)
        self.assertEqual(shown['uzk_registry'],str(self.file))
        updated=apply_settings(
            cfg,registry=source,output=docs,uzk_registry=self.file)
        self.assertEqual(updated['uzk_registry'],str(self.file))
        self.assertTrue(updated['use_shared_uzk'])
        self.assertEqual(updated['documents'],str(forms))
        self.assertEqual(updated['first_passport_documents'],str(docs))
        self.assertEqual(cfg.get('use_shared_uzk'),None)
        self.assertEqual(updated['profiles'],cfg['profiles'])

    def test_missing_xlsm_fails_settings(self):
        source=self.root/'ВЗМК.xlsx';source.write_bytes(b'X')
        forms=self.root/'Templates';forms.mkdir()
        docs=self.root/'Документы';docs.mkdir()
        with self.assertRaisesRegex(Stop,'существующий общий'):
            apply_settings({'documents':str(forms)},
                           registry=source,output=docs,
                           uzk_registry=self.root/'missing.xlsm')

    def test_dated_first_passport_folder_n_s_year_and_root(self):
        root=self.root/'Паспорт';root.mkdir()
        other=root/'11_2026 от 25.10.2026';other.mkdir()
        actual=root/'1_2026 от 09.09.2026';actual.mkdir()
        project=SimpleNamespace(first_number='1/2026',
                                first_date=date(2026,9,9))
        first=_first_roots(root,project)
        self.assertIn(actual,first)
        self.assertNotIn(other,first)
        special=root/'1_1_2026 от 02.08.2026'
        special.mkdir()
        project.first_number='1/1/2026'
        project.first_date=date(2026,8,2)
        self.assertIn(special,_first_roots(root,project))
        self.assertNotIn(actual,_first_roots(root,project))

    def test_output_directory_is_existing_project_passport_folder(self):
        documents=self.root/'Документы';documents.mkdir()
        project=documents/'Гараж казахстан'
        folder=project/'Паспорт';folder.mkdir(parents=True)
        cfg={'output':str(documents),'training':False}
        profile={'first_passport_sources':{'passport_root':str(folder)}}
        self.assertEqual(output_passport_directory(
            cfg,'Гараж Казахстан',profile),folder)
        example=output_passport_directory(
            cfg,'Гараж Казахстан',profile)/'11_2026 от 25.10.2026'
        self.assertEqual(example.parent,folder)
        self.assertFalse(example.exists())
        self.assertEqual(list(folder.iterdir()),[])

    def test_training_output_does_not_write_in_real_archive(self):
        documents=self.root/'Документы';documents.mkdir()
        actual=documents/'Гараж казахстан'/'Паспорт'
        actual.mkdir(parents=True)
        training=self.root/'Training';training.mkdir()
        cfg={'output':str(training),'training':True}
        profile={'first_passport_sources':{'passport_root':str(actual)}}
        dest=output_passport_directory(cfg,'Гараж Казахстан',profile)
        self.assertEqual(dest,training/'Гараж Казахстан'/'Паспорт')
        self.assertFalse(dest.exists())

    def test_wrong_output_root_is_refused_for_real_passport(self):
        documents=self.root/'Документы';documents.mkdir()
        actual=documents/'Гараж казахстан'/'Паспорт'
        actual.mkdir(parents=True)
        different=self.root/'Разное';different.mkdir()
        with self.assertRaisesRegex(Stop,'вне папки объекта'):
            output_passport_directory(
                {'output':str(different),'training':False},
                'Гараж Казахстан',
                {'first_passport_sources':{'passport_root':str(actual)}})

    def test_internal_clean_forms_found_without_manual_profile(self):
        forms=self.root/'Templates'/'Гараж Казахстан';forms.mkdir(parents=True)
        for name in ('ВЗМК_ВИК_ШАБЛОН.docx','ВЗМК_УЗК_ШАБЛОН.docx',
                     'ВЗМК_Документ_о_качестве.docx',
                     'akz0.xlsx','akz_appendix0.xlsx'):
            (forms/name).write_bytes(name.encode())
        result=prepared_project(
            {'documents':str(forms.parent)},'Гараж Казахстан')
        self.assertEqual(
            set(result['templates']),
            {'vik','uzk','quality','akz0','akz_appendix0'})
        self.assertTrue(all(Path(item['path']).parent==forms
                            for item in result['templates'].values()))
        self.assertEqual(result['forms_origin'],
                         'AUTO_DISCOVERED_PREPARED_INTERNAL')

    def test_no_internal_forms_never_reuses_filled_first_passport(self):
        forms=self.root/'Templates'/'Гараж Казахстан';forms.mkdir(parents=True)
        (forms/'1_2026 - ВИК.docx').write_bytes(b'FULLY HISTORICAL ACT')
        with self.assertRaisesRegex(Stop,'Нет очищенных форм'):
            prepared_project(
                {'documents':str(forms.parent)},'Гараж Казахстан')

    def test_daily_ui_no_template_button_and_no_csv(self):
        import inspect
        app=inspect.getsource(gui.App)
        self.assertNotIn("('Шаблоны объекта',self.profile)",app)
        self.assertNotIn('Загрузить CSV УЗК',app)
        self.assertNotIn('Создать CSV-шаблон',app)
        self.assertIn('ОБЩИЙ реестр УЗК.xlsm',app)
        self.assertIn('Сформировать пакет',app)

    def test_check_rejects_shared_source_changed_after_quantity_read(self):
        registry=self.root/'registry.xlsx';registry.write_bytes(b'ISOLATED REGISTRY')
        shipment=self.root/'vom.xlsx';shipment.write_bytes(b'ISOLATED VOM')
        forms=self.root/'forms';forms.mkdir()
        form=forms/'uzk.docx';form.write_bytes(b'ISOLATED FORM')
        output=self.root/'output';output.mkdir()
        cfg={'registry':str(registry),'documents':str(forms),
             'output':str(output),'use_shared_uzk':True,
             'uzk_registry':str(self.file)}
        registry_data=SimpleNamespace(projects={'P':[]},digest=sha(registry))
        shipment_data=SimpleNamespace(path=shipment,digest=sha(shipment),
            rows=(('M','Beam',1,1,1,1,1),),quantity=1,mass=1,area=1)
        facts={'date':'2026-10-25','evidence':[],
               'marks':[{'mark':'M','uzk_quantity':1}]}
        workbook(self.file,[['Заказ','Марка элемента','Кол-во (шт)'],['P','M',1]])
        def control_validation(*args,**kwargs):
            # Simulate a valid concurrent edit after CHECK calculated the
            # quantities, before its immutable input snapshot is created.
            workbook(self.file,[['Заказ','Марка элемента','Кол-во (шт)'],['P','M',9]])
            return []
        with ExitStack() as stack:
            for name,value in [
                ('production_app.core.model.validate_paths',(registry,forms,output)),
                ('production_app.core.pending',[]),
                ('production_app.core.read_registry',registry_data),
                ('production_app.core.model.select',SimpleNamespace(next_number='1/2026')),
                ('production_app.core.model.profile',{'templates':{'uzk':{'path':str(form)}}}),
                ('production_app.core.model.visible_vom',shipment_data),
                ('production_app.templates.validate',None),
                ('production_app.word_controls.inspect_word_fields',(set(),set())),
                ('production_app.real_forms.validate_uzk_details',None),
                ('production_app.real_forms.require_form_inputs',None),
                ('production_app.real_forms.require_akz_form_inputs',None),
                ('production_app.core.local_root',self.root/'local'),
            ]:stack.enter_context(patch(name,return_value=value))
            stack.enter_context(patch('production_app.core.factual.validate',side_effect=control_validation))
            with self.assertRaisesRegex(Stop,'УЗК.*изменился'):
                core.plan(cfg,'P','1/2026',shipment,facts,True)
        self.assertFalse((self.root/'local').exists())


if __name__=='__main__':
    unittest.main(verbosity=2)
