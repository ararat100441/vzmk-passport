from copy import deepcopy
from contextlib import ExitStack
from pathlib import Path
from tempfile import TemporaryDirectory
from types import SimpleNamespace
from unittest.mock import patch
from zipfile import ZipFile
import unittest
from lxml import etree as ET

from production_app.native_acts import W,native_rows,expand_native_uzk,patch_document,text,set_text
from production_app.storage import Stop,sha
from production_app import core
from tools.prepare_first_passport_acts import adapt_uzk,adapt_vik


def fixture(path):
    root=ET.Element(W+'document',nsmap={'w':W[1:-1],'w14':'http://schemas.microsoft.com/office/word/2010/wordml'})
    body=ET.SubElement(root,W+'body')
    for index in range(11):
        if index in (6,8):
            table=ET.SubElement(body,W+'tbl');pr=ET.SubElement(table,W+'tblPr');ET.SubElement(pr,W+'tblW').set(W+'w','10000')
            for row in range(16 if index==6 else 2):
                tr=ET.SubElement(table,W+'tr')
                count=(2 if row<2 else 1 if row<4 else 8) if index==6 else (3 if row==0 else 1)
                for col in range(count):
                    cell=ET.SubElement(tr,W+'tc');cp=ET.SubElement(cell,W+'tcPr');ET.SubElement(cp,W+'tcW').set(W+'w',str(500+col))
                    p=ET.SubElement(cell,W+'p');run=ET.SubElement(p,W+'r');rp=ET.SubElement(run,W+'rPr');ET.SubElement(rp,W+'sz').set(W+'val','20')
                    value='OLD_PERSON годен' if row>=5 else 'HEADER'
                    if index==8:value=('Контроль выполнил','________________','OLD_PERSON')[col] if row==0 else 'OLD_CERTIFICATE'
                    ET.SubElement(run,W+'t').text=value
        elif index==10:ET.SubElement(body,W+'sectPr')
        else:
            p=ET.SubElement(body,W+'p');run=ET.SubElement(p,W+'r');ET.SubElement(run,W+'t').text='Акт № 1/2026 от 04.09.2026' if index==4 else ''
    with ZipFile(path,'w') as z:
        z.writestr('word/document.xml',ET.tostring(root));z.writestr('word/styles.xml',b'PRESERVE_STYLES');z.writestr('opaque.bin',b'PRESERVE_BINARY')


def facts():
    return {'marks':[{'mark':'A','uzk_quantity':8},{'mark':'B','uzk_quantity':0}],
            'uzk_design':[{'mark':'A','profile':'[16П','quantity':8,'length_mm':'1280','thickness_mm':'5','frequency_mhz':'5','angle_deg':'70'},
                          {'mark':'A','profile':'[16П','quantity':8,'length_mm':'1024','thickness_mm':'8.4','frequency_mhz':'5','angle_deg':'70'}],
            'uzk_details':[{'mark':'A','welders':'CURRENT','date':'2026-10-25','frequency_mhz':'5','angle_deg':'70','defects':'CURRENT DEFECTS','conclusion':'CURRENT VERDICT','basis':'CURRENT PROTOCOL'}]}


class NativeActs(unittest.TestCase):
    def setUp(self):self.tmp=TemporaryDirectory();self.root=Path(self.tmp.name)
    def tearDown(self):self.tmp.cleanup()

    def test_draft_contains_design_but_no_implied_pass_or_date(self):
        f=facts();f.pop('uzk_details');row=native_rows(f,draft=True)[0]
        self.assertEqual(row[0],'A');self.assertIn('8 шт.',row[1])
        self.assertIn('L=1280х5 мм',row[1]);self.assertIn('L=1024х8,4 мм',row[1])
        self.assertEqual([row[i] for i in (2,3,6,7)],['НЕТ ДАННЫХ']*4)

    def test_second_vik_act_has_explicit_page_start_after_clearing(self):
        root=ET.Element(W+'document');body=ET.SubElement(root,W+'body')
        sizes={0:(1,1),3:(13,1),5:(2,1),7:(3,1),12:(3,2),13:(1,1),18:(8,1),22:(5,2)}
        for i in range(25):
            if i in sizes:
                table=ET.SubElement(body,W+'tbl');nr,nc=sizes[i]
                for row in range(nr):
                    tr=ET.SubElement(table,W+'tr')
                    for col in range(nc):
                        c=ET.SubElement(tr,W+'tc');p=ET.SubElement(c,W+'p');r=ET.SubElement(p,W+'r')
                        ET.SubElement(r,W+'t').text='Визуального и измерительного контроля' if i==0 else 'CURRENT SLOT'
            else:ET.SubElement(body,W+'sectPr' if i==24 else W+'p')
        second=list(body)[13]
        cell=second.find('.//'+W+'tc');paragraph=cell.find(W+'p');cell.remove(paragraph)
        control=ET.SubElement(cell,W+'sdt');control_props=ET.SubElement(control,W+'sdtPr')
        ET.SubElement(control_props,W+'tag').set(W+'val','PASSPORT_NO')
        content=ET.SubElement(control,W+'sdtContent');content.append(paragraph)
        set_text(paragraph,'1/2026-1')
        adapt_vik(root)
        self.assertEqual(control.find(W+'sdtPr/'+W+'tag').get(W+'val'),'VIK_PROCESS_NUMBER')
        self.assertIn('{{VIK_PROCESS_NUMBER}}',text(control))
        previous=second.getprevious()
        self.assertEqual(previous.tag,W+'p')
        self.assertIsNotNone(previous.find(W+'pPr/'+W+'pageBreakBefore'))
        self.assertIsNotNone(previous.find(W+'pPr/'+W+'keepNext'))

    def test_actual_keeps_only_current_supplied_verdict(self):
        self.assertEqual(native_rows(facts())[0][7],'CURRENT VERDICT')

    def test_missing_actual_record_stops(self):
        f=facts();f['uzk_details']=[]
        with self.assertRaisesRegex(Stop,'текущих результатов'):native_rows(f)

    def test_two_observed_records_are_not_silently_merged(self):
        f=facts();f['uzk_details']*=2
        with self.assertRaises(Stop):native_rows(f)

    def test_actual_angle_is_not_overwritten_by_design(self):
        f=facts();f['uzk_details'][0]['angle_deg']='65'
        with self.assertRaisesRegex(Stop,'angle_deg'):native_rows(f)

    def test_actual_frequency_is_not_overwritten(self):
        f=facts();f['uzk_details'][0]['frequency_mhz']='2.5'
        with self.assertRaisesRegex(Stop,'frequency_mhz'):native_rows(f)

    def test_missing_controlled_mark_cannot_disappear(self):
        f=facts();f['marks'][1]['uzk_quantity']=1
        with self.assertRaisesRegex(Stop,'охватывает'):native_rows(f)

    def test_design_parts_do_not_sum_different_marks(self):
        f=facts();f['marks'][1]['uzk_quantity']=8
        for row in deepcopy(f['uzk_design']):row['mark']='B';f['uzk_design'].append(row)
        rows=native_rows(f,draft=True)
        self.assertEqual([r[0] for r in rows],['A','B'])
        self.assertTrue(all('1280' in r[1] and '2560' not in r[1] for r in rows))

    def test_clearing_preserves_opaque_parts_namespaces_and_source(self):
        source=self.root/'source.docx';fixture(source);before=sha(source)
        cleared=self.root/'cleared.docx';patch_document(source,cleared,adapt_uzk)
        with ZipFile(cleared) as z:
            xml=z.read('word/document.xml');root=ET.fromstring(xml)
            self.assertIn(b'xmlns:w14=',xml)
            self.assertNotIn('OLD_PERSON',text(root));self.assertNotIn('годен',text(root))
            self.assertEqual(z.read('opaque.bin'),b'PRESERVE_BINARY');self.assertEqual(z.read('word/styles.xml'),b'PRESERVE_STYLES')
            tables=root.findall('.//'+W+'tbl');self.assertEqual(len(tables),3)
            self.assertIsNotNone(tables[1].find('.//'+W+'tblHeader'))
        self.assertEqual(sha(source),before)

    def test_expanded_rows_keep_source_widths_and_no_split(self):
        source=self.root/'source.docx';fixture(source)
        cleared=self.root/'cleared.docx';patch_document(source,cleared,adapt_uzk)
        output=self.root/'output.docx';expand_native_uzk(cleared,output,facts())
        with ZipFile(output) as z:
            root=ET.fromstring(z.read('word/document.xml'));table=root.findall('.//'+W+'tbl')[1]
            rows=table.findall(W+'tr');self.assertEqual(len(rows),2)
            self.assertEqual(len(rows[1].findall(W+'tc')),8)
            self.assertIsNotNone(rows[1].find(W+'trPr/'+W+'cantSplit'))
            self.assertEqual([c.find(W+'tcPr/'+W+'tcW').get(W+'w') for c in rows[1].findall(W+'tc')],[str(500+i) for i in range(8)])
            self.assertNotIn('{{UZK_NATIVE_ROW}}',text(root));self.assertNotIn('{{UZK_DESIGN}}',text(root))
            self.assertEqual(z.read('opaque.bin'),b'PRESERVE_BINARY')

    def test_refuse_original_or_existing_output(self):
        source=self.root/'source.docx';fixture(source)
        with self.assertRaises(Stop):patch_document(source,source,adapt_uzk)

    def test_geometry_change_after_calculation_is_rejected_before_job(self):
        reg=self.root/'registry.xlsx';reg.write_bytes(b'REGISTRY')
        shipment=self.root/'vom.xlsx';shipment.write_bytes(b'VOM')
        form=self.root/'uzk.docx';form.write_bytes(b'FORM')
        geometry=self.root/'profiles.xlsx';geometry.write_bytes(b'GEOMETRY_A')
        output=self.root/'output';output.mkdir()
        cfg={'registry':str(reg),'documents':str(self.root),'output':str(output)}
        r=SimpleNamespace(projects={'P':[]},digest=sha(reg))
        v=SimpleNamespace(path=shipment,digest=sha(shipment),mass=1,rows=[('M','N',1,1,1,1,1)])
        f={'date':'2026-10-25','marks':[],'evidence':[]}
        def mutate(*args,**kwargs):geometry.write_bytes(b'GEOMETRY_B');return []
        with ExitStack() as stack:
            for key,value in [('core.model.validate_paths',(reg,self.root,output)),('core.pending',[]),('core.read_registry',r),
                    ('core.model.select',SimpleNamespace(next_number='1/2026')),('core.model.profile',{'templates':{'uzk':{'path':str(form)}}}),
                    ('core.model.visible_vom',v),('word_controls.inspect_word_fields',({'UZK_DESIGN'},set())),
                    ('automatic_acts.find_approved_geometry',geometry),('automatic_acts.calculate_design_characteristics',[]),
                    ('templates.validate',None),('real_forms.validate_uzk_details',None),('real_forms.require_form_inputs',None),('real_forms.require_akz_form_inputs',None)]:
                stack.enter_context(patch('production_app.'+key,return_value=value))
            stack.enter_context(patch('production_app.core.factual.validate',side_effect=mutate))
            with self.assertRaisesRegex(Stop,'геометрии изменился'):core.plan(cfg,'P','1/2026',shipment,f,True)


if __name__=='__main__':unittest.main(verbosity=2)
