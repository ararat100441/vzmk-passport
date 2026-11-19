from pathlib import Path
from tempfile import TemporaryDirectory
import unittest

from production_app.storage import Stop
from production_app.simple_facts import prepare_for_check
from production_app.measurement_ranges import verify_midrange_provenance

def inputs(coats):
    data={'coats':coats, 'air':{'min':'20','max':'26'},
          'surface':{'min':'18','max':'20'}, 'humidity':{'min':'45','max':'60'},
          'layer_intervals':[{'layer':i,'min':'80','max':'100'} for i in range(1,coats+1)]}
    if coats==1:
        data['wet']={'min':'120','max':'140'}
    else:
        data['wet_intervals']=[{'layer':i,'min':'120','max':'140'} for i in range(1,coats+1)]
    return data

class FactsProjectTests(unittest.TestCase):
    def test_project_isolation_and_reopen(self):
        from production_app.fact_defaults import load_defaults, remember_defaults, seed_facts
        from production_app.draft_values import preview_values
        with TemporaryDirectory() as d:
            data={'marks':[{'mark':'M','profile':'Труба 20','akz':0,'standard':'stored',
                            'uzk_quantity':9,'uzk_result':'ГОДЕН'}],
                  'akz_drafts':{'0':inputs(1)},'confirmed':True}
            data['akz_preview']={'0':preview_values(inputs(1),{},'draft-a')}
            remember_defaults('Миндаль',False,'draft-a',data,d)
            data['marks'][0]['profile']='Швеллер 16'
            remember_defaults('Гараж Казахстан',False,'draft-b',data,d)
            m=seed_facts(load_defaults('Миндаль',False,d),'draft-a')
            g=seed_facts(load_defaults('Гараж Казахстан',False,d),'draft-b')
            self.assertEqual(m['marks'][0]['profile'],'Труба 20')
            self.assertEqual(g['marks'][0]['profile'],'Швеллер 16')
            self.assertEqual(m['akz_preview'],data['akz_preview'])
            self.assertNotIn('akz_unconfirmed',m)
            self.assertNotIn('confirmed',m)
            self.assertNotIn('uzk_quantity',m['marks'][0])
            self.assertNotIn('uzk_result',m['marks'][0])
            self.assertFalse(seed_facts(load_defaults('Новый',False,d),'draft-a')['marks'])
            self.assertFalse(seed_facts(load_defaults('Миндаль',True,d),'draft-a')['marks'])

    def test_random_values_cached_and_layers_independent(self):
        from decimal import Decimal
        from production_app.draft_values import draft_climate, preview_values
        for coats in (1,2,3):
            data=inputs(coats)
            samples=preview_values(data,{},'shipment')
            self.assertEqual(samples,preview_values(data,samples,'another-seed'))
            self.assertTrue(Decimal('20')<=Decimal(samples['air']['value'])<=Decimal('26'))
            self.assertEqual(Decimal(samples['air']['value']).as_tuple().exponent,-1)
            self.assertTrue(45<=int(samples['humidity']['value'])<=60)
            for i in range(1,coats+1):
                self.assertTrue(80<=int(samples[f'layer_{i}']['value'])<=100)
                self.assertTrue(120<=int(samples['wet' if coats==1 else f'wet_{i}']['value'])<=140)
            climate=draft_climate(samples,coats)
            self.assertEqual(Decimal(climate['thickness']),
                             sum(Decimal(samples[f'layer_{i}']['value']) for i in range(1,coats+1)))
            if coats>1:
                self.assertEqual(len(climate['dry_averages']),coats)
                self.assertEqual(len(climate['wet_averages']),coats)

    def test_samples_respect_equal_or_narrow_intervals_and_accept_zero(self):
        from decimal import Decimal
        from production_app.draft_values import sample_interval
        for key, minimum, maximum in (('air', '-0.09', '-0.01'), ('air', 0, 0),
                                      ('wet', '120,12', '120,14'), ('humidity', 100, 100)):
            value=sample_interval(key,{'min':minimum,'max':maximum},None,'draft')
            self.assertTrue(Decimal(str(minimum).replace(',','.')) <= Decimal(value['value'])
                            <= Decimal(str(maximum).replace(',','.')))
        self.assertIsNone(sample_interval('wet',{'min':'120','max':''},None,'draft'))

    def test_prefill_passes_measurement_pipeline_for_all_variants(self):
        for variant, coats in enumerate((1, 1, 2, 3)):
            with self.subTest(variant=variant):
                key=str(variant)
                data={'simple_ui':True,'marks':[{'mark':'M','akz':variant}],
                      'akz_drafts':{key:inputs(coats)},'akz_unconfirmed':{key:['air.min']},
                      'akz_preview':{key:{'air':{'value':'26'}}}}
                current=prepare_for_check(data)
                self.assertNotIn('akz_preview',current)
                climate,=current['climates']
                self.assertEqual(climate['air'],'23')
                self.assertEqual(climate['thickness'],str(90*coats))
                verify_midrange_provenance(climate)
                for invalid in ({'min':'26','max':'20'}, {'min':'20','max':''},
                                {'min':'20','max':'999'}, {'min':'wrong','max':'26'}):
                    data['akz_drafts'][key]['air']=invalid
                    with self.assertRaises(Stop):
                        prepare_for_check(data)
        with self.assertRaises(Stop):
            verify_midrange_provenance({'midrange_source':'DRAFT_RANGE_SAMPLE'})
