"""AKZ wet and dry per-coat GUI range logic; synthetic tests only."""
import unittest
from copy import deepcopy

from production_app.facts import validate_layer_ranges, validate_dry_averages
from production_app.measurement_ranges import measurements_from_ranges, verify_midrange_provenance
from production_app.akz_layer_totals import calculate_wet_layers, with_wet_layers
from production_app.simple_facts import prepare_for_check
from production_app.office import akz_values
from production_app.storage import Stop


def source_inputs(coats):
    dry=[('132','136'),('131','135'),('71','75')]
    wet=[('156','158'),('152','154'),('84','86')]
    data={'air':{'min':'24','max':'26'},
          'surface':{'min':'20','max':'22'},
          'humidity':{'min':'45','max':'55'},
          'layer_intervals':[
              {'layer':i+1,'min':dry[i][0],'max':dry[i][1]}
              for i in range(coats)]}
    if coats==1:
        data['wet']={'min':'156','max':'158'}
    else:
        data['wet_intervals']=[
            {'layer':i+1,'min':wet[i][0],'max':wet[i][1]}
            for i in range(coats)]
    return data


def current_draft(coats):
    return {'simple_ui':True,
            'marks':[{'mark':'TEST-MARK','akz':0,'paint':'TEST ONLY'}],
            'akz_drafts':{'0':{'coats':coats,**source_inputs(coats)}}}


class PerLayerWetDry(unittest.TestCase):
    def test_one_coat_existing_values_and_verification_unchanged(self):
        result=measurements_from_ranges(0,source_inputs(1),coats=1)
        self.assertEqual((result['wet_thickness'],result['thickness']),
                         ('157','134'))
        self.assertTrue(verify_midrange_provenance(result))

    def test_two_coats_support_independent_wet_and_dry_input(self):
        result=prepare_for_check(current_draft(2))['climates'][0]
        self.assertEqual([x['average'] for x in result['dry_averages']],['134','267'])
        self.assertEqual([x['average'] for x in result['wet_averages']],['157','153'])
        self.assertEqual([x['cumulative'] for x in result['wet_averages']],['157','310'])
        self.assertEqual((result['wet_total'],result['thickness']),('310','267'))
        self.assertEqual(len(validate_layer_ranges(result['layers'])),2)
        self.assertEqual(len(validate_dry_averages(result['layers'],
                                                  result['dry_averages'])),2)
        self.assertTrue(verify_midrange_provenance(result))
        tokens=akz_values(result)
        self.assertEqual((tokens['LAYER_2_DRY_AVG'],tokens['LAYER_2_WET_AVG'],
                          tokens['LAYER_2_WET_TOTAL']),('267','153','310'))

    def test_three_coats_keep_cumulative_reference_dry_values(self):
        result=prepare_for_check(current_draft(3))['climates'][0]
        self.assertEqual([x['average'] for x in result['dry_averages']],
                         ['134','267','340'])
        self.assertEqual([x['average'] for x in result['wet_averages']],
                         ['157','153','85'])
        self.assertEqual([x['cumulative'] for x in result['wet_averages']],
                         ['157','310','395'])
        self.assertEqual(result['thickness'],'340')
        self.assertTrue(verify_midrange_provenance(result))
        tokens=akz_values(result)
        self.assertEqual(tokens['LAYER_3_DRY_AVG'],'340')
        self.assertEqual(tokens['LAYER_3_WET_AVG'],'85')
        self.assertEqual(tokens['LAYER_3_WET_TOTAL'],'395')
        self.assertEqual(tokens['WET_TOTAL'],'395')
        # Current form-specific AKZ_LAYER_* evidence is not invented.
        self.assertNotIn('AKZ_LAYER_3_WET_AVG',tokens)

    def test_reject_missing_empty_reversed_and_duplicate_wet_layers(self):
        for coats in (2,3):
            with self.subTest(coats=coats):
                a=current_draft(coats)
                a['akz_drafts']['0'].pop('wet_intervals')
                with self.assertRaises(Stop):prepare_for_check(a)
                a=current_draft(coats)
                a['akz_drafts']['0']['wet_intervals'][0]['max']=''
                with self.assertRaises(Stop):prepare_for_check(a)
                a=current_draft(coats)
                a['akz_drafts']['0']['wet_intervals'][0]['min']='999'
                with self.assertRaises(Stop):prepare_for_check(a)
                a=current_draft(coats)
                a['akz_drafts']['0']['wet_intervals'][0]['layer']=2
                with self.assertRaises(Stop):prepare_for_check(a)

    def test_tampering_wet_totals_is_detected_before_issuance(self):
        original=prepare_for_check(current_draft(3))['climates'][0]
        for variant in ('total','average','cumulative','range','dry_average'):
            with self.subTest(variant=variant):
                data=deepcopy(original)
                if variant=='dry_average':data['dry_averages'][1]['average']='999'
                if variant=='total':data['wet_total']='999'
                if variant=='average':data['wet_averages'][1]['average']='999'
                if variant=='cumulative':data['wet_averages'][2]['cumulative']='999'
                if variant=='range':data['measurement_intervals']['wet_layers'][0]['min']='999'
                with self.assertRaises(Stop):verify_midrange_provenance(data)

    def test_three_coat_legacy_low_level_calculation_remains_supported(self):
        data=source_inputs(3)
        data.pop('wet_intervals')
        old=measurements_from_ranges(0,data,coats=3)
        self.assertEqual(old['thickness'],'340')
        self.assertNotIn('wet_total',old)
        self.assertTrue(verify_midrange_provenance(old))

    def test_wet_ranges_accept_comma_decimal_separator_and_order(self):
        wet=[{'layer':2,'min':'150,5','max':'160,5'},
             {'layer':1,'min':'155','max':'159'}]
        rows,averages,total=calculate_wet_layers(wet,2)
        self.assertEqual([x['layer'] for x in rows],[1,2])
        self.assertEqual([x['average'] for x in averages],['157','155.5'])
        self.assertEqual(total,'312.5')
        climate=measurements_from_ranges(0,source_inputs(2),coats=2)
        combined=with_wet_layers(climate,wet,2)
        self.assertEqual(combined['thickness'],'267')
        self.assertTrue(verify_midrange_provenance(combined))


if __name__=='__main__':
    unittest.main()
