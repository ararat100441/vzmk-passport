"""Project-local АКЗ defaults. Synthetic scenarios only; no real registry."""
import json
import tempfile
import unittest
from pathlib import Path

from production_app.akz_presets import last_for_project, remember
from production_app.facts import validate_layer_ranges, validate_dry_averages
from production_app.office import akz_values
from production_app.storage import Stop

CLIMATE = {'variant': 0, 'air': '24.5', 'surface': '27.3',
           'humidity': '54', 'dew': '14.8', 'thickness': '335',
           'material': 'SYNTHETIC', 'basis': 'TEST ONLY'}
LAYERS = [{'layer': n, 'min': str(70 + n * 10),
           'max': str(100 + n * 20)} for n in (1, 2, 3)]
DRY_AVERAGES = [{'layer':1,'average':'100'},
                {'layer':2,'average':'210'},
                {'layer':3,'average':'350'}]


class Presets(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.folder = Path(self.tmp.name)

    def tearDown(self):
        self.tmp.cleanup()

    def test_missing_project_starts_blank(self):
        self.assertEqual(last_for_project('P-NEW', self.folder)['variants'], {})
        self.assertFalse((self.folder / 'akz_presets.json').exists())

    def test_saved_per_project_with_three_ranges(self):
        item = dict(CLIMATE, layers=LAYERS)
        self.assertTrue(remember('P-A', [item], '2026-10-19', self.folder))
        result = last_for_project('P-A', self.folder)
        self.assertEqual(result['variants']['0']['climate']['air'], '24.5')
        self.assertEqual(result['variants']['0']['layers'], LAYERS)
        self.assertEqual(result['variants']['0']['last_akz_date'], '2026-10-19')
        self.assertEqual(last_for_project('P-B', self.folder)['variants'], {})

    def test_project_latest_fallback_available_to_new_variant(self):
        remember('P-A', [dict(CLIMATE, layers=LAYERS)], '2026-10-19', self.folder)
        loaded = last_for_project('P-A', self.folder)
        self.assertEqual(loaded['latest']['climate']['air'], '24.5')
        self.assertEqual(loaded['latest']['layers'], LAYERS)
        self.assertNotIn('1', loaded['variants'])
        self.assertEqual(last_for_project('P-B', self.folder)['latest'], {})

    def test_other_project_does_not_mutate_first(self):
        remember('P-A', [dict(CLIMATE, layers=LAYERS)], '2026-10-19', self.folder)
        other = dict(CLIMATE, air='19.2', layers=[
            {'layer': n, 'min': '100', 'max': '150'} for n in (1, 2, 3)])
        remember('P-B', [other], '2026-10-20', self.folder)
        self.assertEqual(last_for_project('P-A', self.folder)['variants']['0']['climate']['air'], '24.5')
        self.assertEqual(last_for_project('P-B', self.folder)['variants']['0']['climate']['air'], '19.2')

    def test_only_climate_and_ranges_saved_not_actual_facts(self):
        remember('P-A', [dict(CLIMATE, layers=LAYERS, dry_averages=DRY_AVERAGES)], '2026-10-19', self.folder)
        raw = (self.folder / 'akz_presets.json').read_text('utf8')
        for forbidden in ('SYNTHETIC', 'TEST ONLY', '"thickness"', '"dry_averages"', '"basis"',
                          '"material"', '"confirmed"', '"responsible"'):
            self.assertNotIn(forbidden, raw)
        self.assertIn('"air"', raw)

    def test_latest_values_override_only_same_project(self):
        remember('P-A', [dict(CLIMATE, layers=LAYERS)], '2026-10-19', self.folder)
        new = dict(CLIMATE, air='18.7', layers=[
            {'layer': n, 'min': '80', 'max': '110'} for n in (1, 2, 3)])
        remember('P-A', [new], '2026-10-20', self.folder)
        saved = last_for_project('P-A', self.folder)['variants']['0']
        self.assertEqual(saved['climate']['air'], '18.7')
        self.assertEqual(saved['layers'][0]['min'], '80')

    def test_missing_ranges_in_latest_issue_clears_old_ranges(self):
        remember('P-A', [dict(CLIMATE, layers=LAYERS)], '2026-10-19', self.folder)
        remember('P-A', [dict(CLIMATE, air='22.0')], '2026-10-20', self.folder)
        saved = last_for_project('P-A', self.folder)['variants']['0']
        self.assertNotIn('layers', saved)
        self.assertEqual(saved['climate']['air'], '22.0')

    def test_other_variants_of_same_project_stay_separate(self):
        item0 = dict(CLIMATE, layers=LAYERS)
        item1 = dict(CLIMATE, variant=1, air='22', layers=[
            {'layer': n, 'min': '50', 'max': '85'} for n in (1, 2, 3)])
        remember('P-A', [item0, item1], '2026-10-19', self.folder)
        saved = last_for_project('P-A', self.folder)['variants']
        self.assertEqual(saved['0']['climate']['air'], '24.5')
        self.assertEqual(saved['1']['climate']['air'], '22')
        self.assertNotEqual(saved['0']['layers'], saved['1']['layers'])

    def test_partial_layer_ranges_refused_without_cache_mutation(self):
        remember('P-A', [dict(CLIMATE, layers=LAYERS)], '2026-10-19', self.folder)
        before = (self.folder / 'akz_presets.json').read_bytes()
        with self.assertRaises(Stop):
            remember('P-A', [dict(CLIMATE, layers=LAYERS[:1])], '2026-10-20', self.folder)
        self.assertEqual((self.folder / 'akz_presets.json').read_bytes(), before)

    def test_invalid_range_refused(self):
        with self.assertRaisesRegex(Stop, 'Минимальная'):
            validate_layer_ranges([
                {'layer': 1, 'min': '150', 'max': '120'},
                *LAYERS[1:]])
        with self.assertRaisesRegex(Stop, 'уникальные'):
            validate_layer_ranges([LAYERS[0], LAYERS[0], LAYERS[1]])

    def test_bad_schema_refused_without_overwrite(self):
        target = self.folder / 'akz_presets.json'
        target.write_text(json.dumps({'schema': 'OTHER', 'projects': {}}))
        with self.assertRaisesRegex(Stop, 'формат'):
            remember('P-A', [dict(CLIMATE, layers=LAYERS)], '2026-10-19', self.folder)
        self.assertIn('OTHER', target.read_text())

    def test_excel_tokens_share_one_daily_climate(self):
        fields = akz_values(dict(CLIMATE, layers=LAYERS, dry_averages=DRY_AVERAGES))
        self.assertEqual(fields['AIR'], '24.5')
        self.assertEqual(fields['LAYER_1_RANGE'], '80–120')
        self.assertEqual(fields['LAYER_2_RANGE'], '90–140')
        self.assertEqual(fields['LAYER_3_RANGE'], '100–160')
        self.assertEqual(fields['LAYER_1_DRY_AVG'], '100')
        self.assertEqual(fields['LAYER_2_DRY_AVG'], '210')
        self.assertEqual(fields['LAYER_3_DRY_AVG'], '350')
        self.assertEqual(fields['LAYER_3_TOTAL_MIN'], '270')
        self.assertEqual(fields['LAYER_3_TOTAL_MAX'], '420')
        self.assertNotIn('LAYERS', fields)

    def test_full_training_fact_validation_accepts_three_ranges(self):
        from decimal import Decimal
        from types import SimpleNamespace
        from production_app.facts import validate
        from production_app.storage import sha
        proof = self.folder / 'example.txt'
        proof.write_text('SYNTHETIC TRAINING ONLY', encoding='utf8')
        facts = {
            'project':'P-A','number':'1/2026','kind':'TRAINING',
            'confirmed':True,'authorized':True,'forms_approved':True,
            'responsible':'SYNTHETIC',
            'start_date':'2026-10-18','date':'2026-10-19',
            'vik_date':'2026-10-18','uzk_date':'2026-10-18','akz_date':'2026-10-19',
            'marks':[{'mark':'M1','profile':'','paint':'SYNTHETIC',
                      'akz':0,'uzk_quantity':'0','uzk_result':'НЕ ТРЕБУЕТСЯ',
                      'basis':'SYNTHETIC BASIS'}],
            'climates':[dict(CLIMATE,surface='22.3',layers=[dict(row) for row in LAYERS],dry_averages=[dict(row) for row in DRY_AVERAGES],material='SYNTHETIC',basis='SYNTHETIC BASIS')],
            'evidence':[{'path':str(proof),'sha256':sha(proof)}]
        }
        vom = SimpleNamespace(rows=(('M1','Part',Decimal(1),Decimal(1),
                                      Decimal(1),Decimal(1),Decimal(1)),))
        self.assertEqual(validate(facts,vom,'P-A','1/2026',training=True),[0])
        facts['climates'][0]['layers'][0]['max']='50'
        with self.assertRaises(Stop):
            validate(facts,vom,'P-A','1/2026',training=True)

    def test_measured_dry_averages_checked_against_cumulative_ranges(self):
        values = validate_dry_averages(LAYERS, DRY_AVERAGES)
        self.assertEqual([row['average'] for row in values], ['100','210','350'])
        self.assertEqual([row['required_total_min'] for row in values], ['80','170','270'])

    def test_average_outside_range_fails_even_if_prior_layers_passed(self):
        bad = [dict(row) for row in DRY_AVERAGES]
        bad[1]['average'] = '270'  # allowed cumulative upper limit is 260
        with self.assertRaisesRegex(Stop, 'выходит за суммарный диапазон'):
            validate_dry_averages(LAYERS, bad)

    def test_partial_averages_refused_not_replaced_by_range_midpoints(self):
        with self.assertRaisesRegex(Stop, 'после каждого'):
            validate_dry_averages(LAYERS, DRY_AVERAGES[:2])
        with self.assertRaisesRegex(Stop, 'после каждого'):
            akz_values(dict(CLIMATE, layers=LAYERS))

    def test_old_template_without_ranges_remains_supported(self):
        fields = akz_values(CLIMATE)
        self.assertIn('THICKNESS', fields)
        self.assertNotIn('LAYER_1_RANGE', fields)


if __name__ == '__main__':
    unittest.main(verbosity=2)
