from contextlib import nullcontext
import sys
import unittest
from unittest.mock import Mock, patch

from passport_app import excel
from production_app import office


class ExcelValidationTests(unittest.TestCase):
    def verify(self, value, formula, has_formula=False, probe_error=None, header=''):
        used=Mock(Value2=value, Formula=formula, HasFormula=has_formula)
        used.SpecialCells.side_effect=probe_error or LookupError('no matching cells')
        headers={key:'' for key in ('LeftHeader','CenterHeader','RightHeader','LeftFooter','CenterFooter','RightFooter')}
        headers['LeftHeader']=header
        sheet=Mock(UsedRange=used, PageSetup=Mock(**headers))
        book=Mock(Worksheets=[sheet])
        with (patch.object(office,'excel_session',side_effect=lambda:nullcontext(Mock())),
              patch.object(office,'word_session',side_effect=lambda:nullcontext(Mock())),
              patch.object(office,'open_excel',return_value=book)):
            return office.verify_documents([{'path':'act.xlsx','role':'akz0'}],'2/2026',draft=False)

    def test_scalar_number_and_empty_sheet_can_be_verified(self):
        for scalar in [42,None]:
            with self.subTest(scalar=scalar):
                self.assertEqual(self.verify(scalar,scalar)[0]['status'],'PASS')

    def test_scalar_text_and_matching_formula_are_checked_whole(self):
        for value,formula in [('{{NUMBER}}','{{NUMBER}}'),('SAMPLE_ONLY','SAMPLE_ONLY'),(17,'=RANDBETWEEN(1,100)')]:
            with self.subTest(value=value), self.assertRaises(office.Stop):
                self.verify(value,formula)

    def test_failed_native_formula_inspection_blocks_document(self):
        with self.assertRaisesRegex(OSError,'RPC unavailable'):
            self.verify(((5,17),),((5,'=A1'),),None,OSError('RPC unavailable'))

    def test_failed_native_formula_inspection_blocks_registry_candidate(self):
        sheet=Mock()
        sheet.Range.return_value.HasFormula=None
        sheet.Range.return_value.SpecialCells.side_effect=OSError('RPC unavailable')
        with self.assertRaisesRegex(OSError,'RPC unavailable'):
            excel.assert_no_formula_errors(sheet,2,8)

    def test_readback_can_inspect_readonly_document(self):
        app=Mock()
        book=app.Workbooks.Open.return_value
        book.ReadOnly=True
        book.Worksheets=[]
        with (patch.object(office,'excel_session',side_effect=lambda:nullcontext(app)),
              patch.object(office,'word_session',side_effect=lambda:nullcontext(Mock()))):
            result=office.verify_documents([{'path':'act.xlsx','role':'akz0'}],'2/2026',draft=False)
        self.assertEqual(result[0]['status'],'PASS')
        self.assertTrue(app.Workbooks.Open.call_args.kwargs['ReadOnly'])
        book.Close.assert_called_once_with(SaveChanges=False)
        with self.assertRaises(excel.PassportError):
            excel.open_excel(app,'candidate.xlsx')

    def test_excel_header_with_unfilled_token_blocks_verification(self):
        with self.assertRaises(office.Stop):
            self.verify((('ok','2/2026'),),(('ok','2/2026'),),header='{{NUMBER}}')

    def test_word_footer_with_unconfirmed_control_blocks_verification(self):
        footer=Mock(Exists=True)
        footer.Range.Text='НЕ ПОДТВЕРЖДЕНО'
        doc=Mock(Sections=[Mock(Headers=[],Footers=[footer])])
        doc.Content.Text='Документ о качестве №2/2026'
        with (patch.object(office,'excel_session',side_effect=lambda:nullcontext(Mock())),
              patch.object(office,'word_session',side_effect=lambda:nullcontext(Mock())),
              patch.object(office,'open_word',return_value=doc),
              self.assertRaises(office.Stop)):
            office.verify_documents([{'path':'act.docx','role':'quality'}],'2/2026',draft=False)

    def test_only_formula_areas_are_examined_and_invalid_counts_fail(self):
        sheet,region=Mock(),Mock(HasFormula=None)
        region.SpecialCells.return_value.Areas=[Mock(Address='$C$1'),Mock(Address='$D$2:$D$3')]
        # Exercise the same dispatch contract in source-only runs without pywin32.
        dispatch = sheet._oleobj_ = Mock()
        dispatch.GetIDsOfNames.return_value = 1
        dispatch.Invoke.side_effect = [0, 0.0]
        with patch.dict(sys.modules, {'pythoncom': Mock(DISPATCH_METHOD=1)}):
            excel.assert_range_formulas(sheet, region)
            dispatch.GetIDsOfNames.assert_called_once_with(1033, 'Evaluate')
            self.assertEqual([call.args for call in dispatch.Invoke.call_args_list], [
                (1, 1033, 1, True, 'SUMPRODUCT(--ISERROR($C$1))'),
                (1, 1033, 1, True, 'SUMPRODUCT(--ISERROR($D$2:$D$3))'),
            ])
            for bad in [1, None, -2146826281, False]:
                with self.subTest(bad=bad), self.assertRaises(excel.PassportError):
                    dispatch.Invoke.side_effect = [0, bad]
                    excel.assert_range_formulas(sheet, region)


@unittest.skipUnless(sys.platform=='win32','requires Microsoft Excel and pywin32')
class NativeExcelValidationTests(unittest.TestCase):
    def test_formula_results_checked_only_in_requested_range(self):
        with excel.excel_session() as app:
            book=app.Workbooks.Add()
            try:
                s=book.Worksheets(1)
                s.Range('A1').Value2=4
                s.Range('B1').Formula='=2+2'
                s.Range('C1').Formula='=1/0'
                s.Range('A2').Formula='=NA()'
                app.CalculateFullRebuild()
                excel.assert_no_formula_errors(s,1,2)
                with self.assertRaises(excel.PassportError):
                    excel.assert_no_formula_errors(s,1,3)
                with self.assertRaises(excel.PassportError):
                    excel.assert_no_formula_errors(s,2,1)
            finally:
                book.Close(SaveChanges=False)
