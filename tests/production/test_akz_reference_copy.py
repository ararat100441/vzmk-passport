"""Current ranges into a copy of the existing act; no historical facts imported."""
from copy import deepcopy
from unittest import TestCase
from unittest.mock import patch

from test_akz_from_first import Sheet,AKZ_FORMULAS
from test_gpt6_akz_wet_dry_layers import current_draft
from production_app.simple_facts import prepare_for_check
from production_app.akz_from_first import fill_first_akz,is_first_akz,reference_coats
from production_app.storage import Stop


FORMULAS={'G8':'=E4','I8':'=H4','D49':'=H4',
          'E22':'=H22-1','H22':'=J22-1','J22':'=H4-1',
          'E30':'=E22','H30':'=H22','J30':'=J22',
          'H27':'=H28*100/L27','J27':'=J28*100/L27'}


class ReferenceCopy(TestCase):
    def sheet(self):
        s=Sheet();s.data.update(E4='1/1/2026 от ',H4=46000,L27=86)
        for i,col in enumerate(('E','H','J'),1):
            s.data[col+'17']=i
            for row in (27,28,31,32,33,34):s.data[col+str(row)]=1
        for cell,formula in FORMULAS.items():s.Range(cell).Formula=formula
        return s

    def climate(self):return prepare_for_check(current_draft(3))['climates'][0]

    def test_existing_series_three_coat_shape_is_recognized(self):
        s=self.sheet();cells={k:{'value':str(v)} for k,v in s.data.items()}
        for k,v in FORMULAS.items():cells[k]={'formula':v[1:]}
        with patch('production_app.akz_from_first._cells',return_value=cells):
            self.assertTrue(is_first_akz('reference.xlsx'))
            self.assertEqual(reference_coats('reference.xlsx'),3)

    def test_native_excel_float_layer_numbers_are_recognized(self):
        s=self.sheet()
        for i,col in enumerate(('E','H','J'),1):s.data[col+'17']=float(i)
        result=fill_first_akz(s,number='12/1/2026',akz_date='2026-11-04',climate=self.climate())
        self.assertEqual(result['coats'],3)

    def test_only_number_date_climate_and_thickness_values_change(self):
        s=self.sheet();s.data.update(E18='REFERENCE PAINT',E19='OLD BATCH',B43='OLD TEXT')
        before=deepcopy(s.data);result=fill_first_akz(s,number='12/1/2026',akz_date='2026-11-04',climate=self.climate())
        self.assertEqual(s.data['E4'],'12/1/2026 от ')
        self.assertEqual([s.data[c+'27'] for c in ('E','H','J')],[157,153,85])
        self.assertEqual([s.data[c+'28'] for c in ('E','H','J')],[134,267,340])
        for cell in set(before)-set(result['updated_cells']):self.assertEqual(s.data[cell],before[cell])
        self.assertTrue(result['unchanged_reference_fields_are_not_current_control'])
        for ref in s.refs.values():self.assertIsNone(ref.Interior.Pattern)

    def test_date_and_nonmeasurement_formulas_are_unchanged(self):
        s=self.sheet();fill_first_akz(s,number='12/1/2026',akz_date='2026-11-04',climate=self.climate())
        for cell,formula in FORMULAS.items():
            if cell not in ('H27','J27'):self.assertEqual(s.Range(cell).Formula,formula)

    def test_invalid_ranges_stop_before_any_write(self):
        s=self.sheet();before=deepcopy(s.data);c=self.climate();c['wet_averages'][1]['average']='999'
        with self.assertRaises(Stop):fill_first_akz(s,number='12/1/2026',akz_date='2026-11-04',climate=c)
        self.assertEqual(s.data,before)

    def test_wrong_coat_count_stops_before_any_write(self):
        s=self.sheet();before=deepcopy(s.data)
        c=prepare_for_check(current_draft(2))['climates'][0]
        with self.assertRaisesRegex(Stop,'число измеренных слоёв'):
            fill_first_akz(s,number='12/1/2026',akz_date='2026-11-04',climate=c)
        self.assertEqual(s.data,before)

    def test_current_one_coat_input_cannot_reuse_old_wet_value(self):
        s=Sheet();s.data.update(E4='1/1/2026 от ',E27=157)
        for cell,formula in AKZ_FORMULAS.items():s.Range(cell).Formula=formula
        draft=current_draft(1);draft['akz_drafts']['0'].pop('wet')
        c=prepare_for_check(draft)['climates'][0];before=deepcopy(s.data)
        with self.assertRaisesRegex(Stop,'нет текущего измерения мокрой'):
            fill_first_akz(s,number='12/1/2026',akz_date='2026-11-04',climate=c)
        self.assertEqual(s.data,before)

    def test_missing_wet_measurements_are_not_copied_from_reference(self):
        s=self.sheet();before=deepcopy(s.data);c=self.climate()
        for key in ('wet_averages','wet_total'):c.pop(key)
        c['measurement_intervals'].pop('wet_layers')
        with self.assertRaisesRegex(Stop,'мокрые измерения'):
            fill_first_akz(s,number='12/1/2026',akz_date='2026-11-04',climate=c)
        self.assertEqual(s.data,before)

    def test_unrecognized_wet_formula_is_not_overwritten(self):
        s=self.sheet();s.Range('H27').Formula='=SUM(A1:A9)';before=deepcopy(s.data)
        with self.assertRaisesRegex(Stop,'исходная формула H27'):
            fill_first_akz(s,number='12/1/2026',akz_date='2026-11-04',climate=self.climate())
        self.assertEqual(s.data,before)
