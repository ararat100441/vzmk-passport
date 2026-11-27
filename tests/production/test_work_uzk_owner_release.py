import unittest
from contextlib import nullcontext
from pathlib import Path
from tempfile import TemporaryDirectory
from types import SimpleNamespace
from unittest.mock import Mock, patch

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt

from production_app import core, office, owner_control
from production_app.storage import Stop, pins, save, sha
from production_app.working_facts import selected_form


class SelectedSummaryForm(unittest.TestCase):
    def test_only_selected_source_with_complete_blank_summary_contract_is_accepted(self):
        with TemporaryDirectory() as tmp:
            path=Path(tmp)/'uzk.docx'
            doc=Document()
            for text in ('Заключение ультразвукового контроля № {{NUMBER}} от {{END_DATE}}',
                         'Договор: {{CONTRACT}}','Заказчик: {{CUSTOMER}}','Объект: {{OBJECT}}',
                         'Шифр: {{CIPHER}}','УЗК: {{UZK_DATE}}','{{UZK_TABLE}}',
                         'Протоколы: {{EVIDENCE}}','Ответственное лицо: {{RESPONSIBLE}}',
                         'Подпись: __________________'):
                doc.add_paragraph(text)
            doc.save(path)
            ref={'path':str(path),'sha256':sha(path)}
            profile={'first_passport_sources':{'first_number':'1/2026','sources':{'uzk':ref}},
                     'templates':{'uzk':{'path':'unselected.docx'}}}
            self.assertEqual(selected_form(profile,'uzk'),ref)

            doc.add_paragraph('Ранее дефекты не обнаружены')
            doc.save(path);ref['sha256']=sha(path)
            with self.assertRaisesRegex(Stop,'Другая форма автоматически не подставлена'):
                selected_form(profile,'uzk')


class OwnerSummaryRelease(unittest.TestCase):
    def setUp(self):
        folder=TemporaryDirectory()
        self.addCleanup(folder.cleanup)
        self.root=Path(folder.name)
        self.draft=self.root/'draft.docx'
        self.output=self.root/'release'
        self.output.mkdir()
        doc=Document()
        doc.add_paragraph(office.PREVIEW_MARK)
        doc.add_paragraph('УЗК № 4/2026')
        rows=[('Марка','ВОМ, шт','УЗК, шт','Результат','Основание','Расчётная геометрия'),
              ('M','2','2','Ожидает подтверждения владельца','','L=75/S=8; L=155/S=6'),
              ('Z','1','0','НЕ ТРЕБУЕТСЯ','План УЗК по текущей отгрузке: 0 шт.','')]
        table=doc.add_table(rows=0,cols=6)
        for row in rows:
            for cell,text in zip(table.add_row().cells,row):
                cell.text=text
        filled,blank=table.rows[1].cells[3:5]
        run=filled.paragraphs[0].runs[0]
        run.font.name='Times New Roman';run.font.size=Pt(11)
        para=blank.paragraphs[0]
        ppr=para._p.get_or_add_pPr()
        marker=OxmlElement('w:rPr')
        fonts=OxmlElement('w:rFonts');fonts.set(qn('w:ascii'),'Times New Roman');fonts.set(qn('w:hAnsi'),'Times New Roman')
        size=OxmlElement('w:sz');size.set(qn('w:val'),'22')
        marker.append(fonts);marker.append(size);ppr.append(marker)
        doc.save(self.draft)
        self.item={'role':'uzk','path':str(self.draft),'owner_uzk_summary':True}

    def verify(self,item,*,draft):
        doc=Document(item['path'])
        text='\n'.join([*(p.text for p in doc.paragraphs),
            *(c.text for table in doc.tables for row in table.rows for c in row.cells)])
        opened=Mock(Content=SimpleNamespace(Text=text),Sections=[])
        with (patch.object(office,'excel_session'),patch.object(office,'word_session'),
              patch.object(office,'open_word',return_value=opened)):
            return office.verify_documents([item],'4/2026',draft=draft)

    def test_draft_validates_and_only_confirmed_release_copy_gets_current_verdict(self):
        original=sha(self.draft)
        self.assertEqual(self.verify(self.item,draft=True)[0]['status'],'PASS')
        with patch.object(office,'excel_session'):
            out,=office.release_copies([self.item],self.output,False,owner_uzk_sha256=original)
        self.assertEqual(sha(self.draft),original)
        report=Document(out['path'])
        rows=[[c.text for c in row.cells] for row in report.tables[0].rows]
        self.assertEqual(rows[1][3],'ГОДЕН')
        self.assertIn('текущем УЗК',rows[1][4])
        self.assertEqual(rows[1][-1],'L=75/S=8; L=155/S=6')
        self.assertEqual(rows[2][3],'НЕ ТРЕБУЕТСЯ')
        self.assertEqual(self.verify(out,draft=False)[0]['status'],'PASS')

    def test_release_preserves_the_result_run_and_blank_basis_style(self):
        with patch.object(office,'excel_session'):
            out,=office.release_copies([self.item],self.output,False,owner_uzk_sha256=sha(self.draft))
        original=Document(self.draft).tables[0].rows[1].cells
        result=Document(out['path']).tables[0].rows[1].cells
        for index in (3,4):
            para=result[index].paragraphs[0]
            run,=para.runs
            self.assertEqual((run.font.name,run.font.size),('Times New Roman',Pt(11)))
            old_para=original[index].paragraphs[0]
            self.assertEqual([item.xml for item in para._p.xpath('./w:pPr')],
                             [item.xml for item in old_para._p.xpath('./w:pPr')])

    def test_pending_summary_cannot_release_without_declaration(self):
        with patch.object(office,'excel_session'):
            with self.assertRaises(Stop):
                office.release_copies([self.item],self.output,False)
        self.assertEqual(list(self.output.iterdir()),[])

    def test_release_rejects_a_copy_with_a_different_verified_hash(self):
        with patch.object(office,'excel_session'):
            with self.assertRaisesRegex(Stop,'проверенн'):
                office.release_copies([self.item],self.output,False,owner_uzk_sha256='different')
        self.assertFalse((self.output/self.draft.name).exists())

    def test_release_rechecks_the_copied_summary_after_excel_before_filling(self):
        original=sha(self.draft)
        with patch.object(office,'excel_session') as excel:
            def later_edit():
                copy=self.output/self.draft.name
                doc=Document(copy)
                doc.tables[0].cell(1,5).text='Changed geometry'
                doc.save(copy)
                return nullcontext()
            excel.side_effect=later_edit
            with self.assertRaisesRegex(Stop,'проверенн'):
                office.release_copies([self.item],self.output,False,owner_uzk_sha256=original)
        self.assertEqual(Document(self.output/self.draft.name).tables[0].cell(1,3).text,
                         'Ожидает подтверждения владельца')

    def test_native_verification_cannot_repin_a_changed_summary(self):
        registry=self.root/'registry.xlsx'
        registry.write_bytes(b'Excel candidate')
        save(self.root/'prepared.json',{'plan_seal':'seal','registry':str(registry),
             'snapshot_sha256':sha(registry),'outputs':[self.item],
             'output_fingerprints':pins([self.draft])})
        plan={'seal':'seal','run':str(self.root),'number':'4/2026','cfg':{}}
        def changed_after_word(*_args):
            self.draft.write_bytes(b'Edited after native verification')
            return [{'file':str(self.draft),'owner_uzk_summary':True,'status':'PASS'}]
        with (patch.object(core,'load_plan',return_value=plan),
              patch.object(office,'verify_documents',side_effect=changed_after_word)):
            with self.assertRaisesRegex(Stop,'изменился'):
                core.verify(self.root/'plan.json')
        self.assertFalse((self.root/'verified.json').exists())

    def test_pending_and_premature_verdicts_are_rejected_in_wrong_phases(self):
        with self.assertRaises(Stop):
            self.verify(self.item,draft=False)
        doc=Document(self.draft)
        doc.tables[0].cell(1,3).text='ГОДЕН'
        doc.save(self.draft)
        with self.assertRaises(Stop):
            self.verify(self.item,draft=True)

    def test_core_requires_approval_for_exact_verified_summary(self):
        facts={'owner_confirmation_required':True,'responsible':'','marks':[]}
        plan={'run':str(self.root),'seal':'seal','facts':facts,'project':'P',
              'number':'4/2026','cfg':{}}
        verified=self.root/'verified.json'
        prepared=self.root/'prepared.json'
        save(prepared,{'outputs':[self.item]})
        save(verified,{'plan_seal':'seal','documents':[{'status':'PASS','file':str(self.draft),
                       'owner_uzk_summary':True}],
                       'owner_uzk_preparation':pins([prepared]),
                       'fingerprints':pins([self.draft])})
        path=self.root/'plan.json'
        self.output.rmdir()
        with (patch.object(core,'load_plan',return_value=plan),
              patch.object(office,'release_copies',side_effect=Stop('release boundary')) as release):
            with self.assertRaisesRegex(Stop,'текущий контроль'):
                core.confirm(path,'',False,approved=True,current_control=False,direct_issue=True)
            self.assertFalse((self.root/'confirmation.json').exists())
            core.confirm(path,'',False,approved=True,current_control=True,direct_issue=True)
            confirmation=core.read(self.root/'confirmation.json')
            self.assertEqual(confirmation['current_control'],owner_control.declaration(plan,sha(verified)))
            before=verified.read_bytes()
            verified.write_bytes(before+b' ')
            with self.assertRaisesRegex(Stop,'Нет подтверждения'):
                core.release(path)
            release.assert_not_called()
            verified.write_bytes(before)
            save(prepared,{'outputs':[self.item|{'owner_uzk_summary':False}]})
            with self.assertRaisesRegex(Stop,'SHA-256 изменился'):
                core.release(path)
            release.assert_not_called()
            with self.assertRaises(Stop):
                core.confirm(path,'',False,approved=True,current_control=True,direct_issue=True)
            save(prepared,{'outputs':[self.item]})
            with self.assertRaisesRegex(Stop,'release boundary'):
                core.release(path)
            release.assert_called_once_with([self.item],self.output,False,owner_uzk_sha256=sha(self.draft))
