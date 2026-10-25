"""VZMK UZK geometry from accepted Stage4 data, not invented inspections.

Synthetic sheets represent the real 16П profile (ГОСТ 8240-97):
h=160, b=64, s=5, t=8.4. VOM/shipment quantities independent per mark.
"""
from pathlib import Path
from tempfile import TemporaryDirectory
from types import SimpleNamespace
import unittest

from openpyxl import Workbook

from production_app.automatic_acts import (
    calculate_design_characteristics,find_approved_geometry,
    joint_text,ordered_vik_names,
)
from production_app.storage import Stop,sha
from production_app.word_controls import inspect_word_fields,BINDABLE_FIELDS
from production_app import office
from zipfile import ZipFile
import inspect


HEADER=[
    'Профиль','Ключ','Тип','НД','h','b','b2','s','t','D','r',
    'Источник','Дата получения','Статус','УЗК правило','Радиус R',
]


def make_book(path,*,duplicate=False,modified=False):
    wb=Workbook()
    sh=wb.active;sh.title='Профили';sh.append(HEADER)
    sh.append([
        '16П','16П','Швеллер','ГОСТ 8240-97',
        160,64,None,5,8.4,None,5,'TEST',None,
        'ПРОВЕРЕНО_ДВУМЯ_ИСТОЧНИКАМИ',
        'ШВЕЛЛЕР_16П_1',8.5,
    ])
    if duplicate:
        sh.append([
            '16П', '16П', 'Швеллер','ДРУГОЙ ГОСТ',
            160,64,None,5,8.4,None,5,'TEST',None,
            'ПРОВЕРЕНО_ДВУМЯ_ИСТОЧНИКАМИ',
            'ШВЕЛЛЕР_16П_1',8.5,
        ])
    if modified:
        sh['B1']='BAD KEY'
    wb.save(path);wb.close()


class AutomaticActs(unittest.TestCase):
    def setUp(self):
        self.temp=TemporaryDirectory()
        self.dir=Path(self.temp.name)
        self.book=self.dir/'Профили.xlsx'
        make_book(self.book)
        self.vom=SimpleNamespace(rows=[
            ('Бп2-1','Балка',5,1,5,1,5),
            ('Бп2-1','Балка',6,1,6,1,6),
            ('Бп2-5','Балка',1,1,1,1,1),
        ])
        self.facts={'marks':[
            {'mark':'Бп2-1','profile':'[16П','uzk_quantity':'8'},
            {'mark':'Бп2-5','profile':'16П','uzk_quantity':'0'},
        ]}
    def tearDown(self):self.temp.cleanup()

    def test_one_mark_two_parts_no_intermark_summation(self):
        before=sha(self.book)
        data=calculate_design_characteristics(self.vom,self.facts,self.book)
        self.assertEqual(len(data),2)
        self.assertEqual([r['mark'] for r in data],['Бп2-1','Бп2-1'])
        by_thick={r['thickness_mm']:r for r in data}
        self.assertEqual(by_thick['5']['length_mm'],'1280')
        self.assertEqual(by_thick['8.4']['length_mm'],'1024')
        self.assertEqual(by_thick['5']['segments'],('стенка',))
        self.assertEqual(by_thick['8.4']['segments'],('полки',))
        self.assertEqual(by_thick['8.4']['quantity'],8)
        self.assertEqual(by_thick['8.4']['frequency_mhz'],'5')
        self.assertEqual(by_thick['8.4']['angle_deg'],'70')
        self.assertEqual(by_thick['5']['geometry_reference']['standard'],
                         'ГОСТ 8240-97')
        self.assertEqual(sha(self.book),before)

    def test_vik_marks_are_new_not_historical(self):
        self.assertEqual(ordered_vik_names(self.vom),
                         ('Бп2-1','Бп2-5'))

    def test_missing_profile_blocks_only_inspected_mark(self):
        self.facts['marks'][0]['profile']=''
        with self.assertRaisesRegex(Stop,'не указан профиль'):
            calculate_design_characteristics(self.vom,self.facts,self.book)

    def test_missing_profile_on_noninspected_mark_does_not_block(self):
        self.facts['marks'][1]['profile']=''
        result=calculate_design_characteristics(self.vom,self.facts,self.book)
        self.assertEqual(len(result),2)

    def test_conflicting_geometry_for_same_profile_stops(self):
        make_book(self.book,duplicate=True)
        with self.assertRaisesRegex(Stop,'единственной'):
            calculate_design_characteristics(self.vom,self.facts,self.book)

    def test_changed_workbook_pin_blocks(self):
        with self.assertRaisesRegex(Stop,'изменился'):
            calculate_design_characteristics(
                self.vom,self.facts,self.book,expected_sha='WRONG')

    def test_no_inspected_marks_yields_empty_without_guesses(self):
        for row in self.facts['marks']:
            row['uzk_quantity']='0'
        self.assertEqual(calculate_design_characteristics(
            self.vom,self.facts,self.book),[])

    def test_no_verdict_or_invented_welder_in_design_data(self):
        data=calculate_design_characteristics(self.vom,self.facts,self.book)
        for item in data:
            for field in ('welders','conclusion','defects','basis','date'):
                self.assertNotIn(field,item)
            self.assertIn('толщина',joint_text(item))
            self.assertIn('длина контроля',joint_text(item))

    def test_missing_accepted_profile_source_refused(self):
        empty=self.dir/'EmptyForms'
        empty.mkdir()
        with self.assertRaisesRegex(Stop,'справочник геометрии'):
            find_approved_geometry(
                {'registry':str(empty/'ВЗМК.xlsx'),
                 'documents':str(empty)},require=True)

    def test_explicit_approved_source_no_new_gui_setting(self):
        path=find_approved_geometry({
            'profiles_workbook':str(self.book),
            'profiles_workbook_sha256':sha(self.book),
        },require=True)
        self.assertEqual(path,self.book)

    def test_bad_sheet_headers_refused(self):
        make_book(self.book,modified=True)
        with self.assertRaisesRegex(Stop,'структура'):
            calculate_design_characteristics(self.vom,self.facts,self.book)

    def test_word_sources_accept_optional_dynamic_mark_and_design_tokens(self):
        with TemporaryDirectory() as tmp:
            file=Path(tmp)/'BlankWord.docx'
            xml=(
                '<?xml version="1.0" encoding="UTF-8"?>'
                '<w:document xmlns:w="http://schemas.openxmlformats.org/'
                'wordprocessingml/2006/main"><w:body>'
                '<w:p><w:r><w:t>{{VIK_MARKS}}</w:t></w:r></w:p>'
                '<w:p><w:r><w:t>{{UZK_DESIGN}}</w:t></w:r></w:p>'
                '</w:body></w:document>'
            )
            with ZipFile(file,'w') as z:
                z.writestr('word/document.xml',xml)
            tokens,controls=inspect_word_fields(file)
            self.assertEqual(tokens,{'VIK_MARKS','UZK_DESIGN'})
            self.assertFalse(controls)
            self.assertTrue(tokens.issubset(BINDABLE_FIELDS))
        code=inspect.getsource(office.render)
        self.assertIn("values['VIK_MARKS']",code)
        self.assertIn("values['UZK_DESIGN']",code)
        self.assertIn('ordered_vik_names(vom)',code)
        self.assertIn("facts.get('uzk_design',[])",code)

    def test_mismatched_vom_marks_refused(self):
        self.facts['marks'].append({
            'mark':'ЧУЖАЯ','profile':'16П','uzk_quantity':'2',
        })
        with self.assertRaisesRegex(Stop,'марки отличаются'):
            calculate_design_characteristics(self.vom,self.facts,self.book)


if __name__=='__main__':
    unittest.main(verbosity=2)
