"""Two-coat E/I original layout: current measurements, no new fact fields."""
from copy import deepcopy
from pathlib import Path
from types import SimpleNamespace
from unittest import TestCase
from unittest.mock import patch,MagicMock
from contextlib import nullcontext
from test_akz_from_first import Sheet
from test_gpt6_akz_wet_dry_layers import current_draft
from production_app.simple_facts import prepare_for_check
from production_app.akz_from_first import fill_first_akz,is_first_akz,reference_coats
from production_app.storage import Stop

FORMULAS={'G8':'=E4','I8':'=H4','D49':'=H4','E22':'=I22-1','I22':'=H4-1',
          'E30':'=E22','I30':'=I22','E27':'=E28*100/L27','I27':'=I28*100/L27'}

class TwoReference(TestCase):
    def sheet(self):
        s=Sheet();s.data.update(E4='1/2025  от ',E17=1.,H17=2.,H4=46000,L27=86)
        for col in ('E','I'):
            for row in (28,31,32,33,34):s.data[col+str(row)]=1
        for cell,f in FORMULAS.items():s.Range(cell).Formula=f
        return s
    def climate(self):return prepare_for_check(current_draft(2))['climates'][0]

    def test_header_h17_is_two_but_measurement_column_is_i(self):
        s=self.sheet();cells={k:{'value':str(int(v)) if isinstance(v,float) and v.is_integer() else str(v)} for k,v in s.data.items()}
        for k,f in FORMULAS.items():cells[k]={'formula':f[1:]}
        with patch('production_app.akz_from_first._cells',return_value=cells):
            self.assertTrue(is_first_akz('original.xlsx'))
            self.assertEqual(reference_coats('original.xlsx'),2)

    def test_variant_three_still_has_two_coats_and_cumulative_dry(self):
        s=self.sheet();c=self.climate();c['variant']=3
        result=fill_first_akz(s,number='12/2026',akz_date='2026-11-04',climate=c)
        self.assertEqual(result['coats'],2)
        self.assertEqual(s.data['E4'],'12/2026-3 от ')
        self.assertEqual((s.data['E27'],s.data['I27']),(157,153))
        self.assertEqual((s.data['E28'],s.data['I28']),(134,267))
        self.assertNotIn('H28',result['updated_cells'])
        self.assertNotIn('J28',result['updated_cells'])

    def test_random_climate_helpers_are_current_values_in_copy(self):
        s=self.sheet();s.data['E19']='REFERENCE BATCH'
        for cell in ('L31','M31','L33','M33'):s.Range(cell).Formula='=E31+RANDBETWEEN(0.1,2)'
        fill_first_akz(s,number='12/2026',akz_date='2026-11-04',climate=self.climate())
        self.assertEqual([s.data[p] for p in ('L31','M31','L33','M33')],[21,21,50,50])
        self.assertEqual(s.data['E19'],'REFERENCE BATCH')
        for cell,f in FORMULAS.items():
            if cell not in ('E27','I27'):self.assertEqual(s.Range(cell).Formula,f)

    def test_wet_thickness_is_not_printed_as_currency(self):
        s=self.sheet();s.Range('E27').NumberFormat='General';s.Range('I27').NumberFormat='0 ₽'
        fill_first_akz(s,number='12/2026',akz_date='2026-11-04',climate=self.climate())
        self.assertEqual(s.Range('I27').NumberFormat,'General')

    def test_changed_midrange_stops_before_any_write(self):
        s=self.sheet();before=deepcopy(s.data);c=self.climate();c['dry_averages'][1]['average']='401'
        with self.assertRaises(Stop):fill_first_akz(s,number='12/2026',akz_date='2026-11-04',climate=c)
        self.assertEqual(s.data,before)

    def test_three_coat_measurements_do_not_fit_two_coat_source(self):
        s=self.sheet();before=deepcopy(s.data);c=prepare_for_check(current_draft(3))['climates'][0]
        with self.assertRaisesRegex(Stop,'число измеренных слоёв'):
            fill_first_akz(s,number='12/2026',akz_date='2026-11-04',climate=c)
        self.assertEqual(s.data,before)

    def test_unknown_formula_is_not_overwritten(self):
        s=self.sheet();before=deepcopy(s.data);s.Range('I27').Formula='=SUM(A1:A9)'
        with self.assertRaisesRegex(Stop,'исходная формула I27'):
            fill_first_akz(s,number='12/2026',akz_date='2026-11-04',climate=self.climate())
        self.assertEqual(s.data,before)

    def test_pdf_export_uses_act_sheet_without_saving_auxiliary(self):
        from production_app.print_pdf import _export_excel
        book=MagicMock();book.ReadOnly=True
        excel=MagicMock();excel.Workbooks.Open.return_value=book
        item={'role':'akz3','source':Path('source.xlsx'),'pdf':Path('output.pdf')}
        with patch('production_app.print_pdf.excel_session',return_value=nullcontext(excel)), \
             patch('production_app.akz_from_first.is_first_akz',return_value=True), \
             patch('production_app.akz_from_first.reference_coats',return_value=2):
            _export_excel([item])
        book.Worksheets.assert_called_once_with(1)
        book.Worksheets.return_value.ExportAsFixedFormat.assert_called_once()
        book.ExportAsFixedFormat.assert_not_called();book.Save.assert_not_called()
        book.Close.assert_called_once_with(SaveChanges=False)
