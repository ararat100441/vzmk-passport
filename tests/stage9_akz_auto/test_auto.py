"""New Stage 9 contracts. Native/frozen integration is recorded separately."""
import copy,json,math,os,sys,tempfile,types,unittest
from pathlib import Path
from uuid import uuid4
from unittest.mock import patch
ROOT=Path(__file__).resolve().parents[2];sys.path.insert(0,str(ROOT))
from stage9_app import akz_auto as A,core
from stage9_app.workspace import save
from stage8_closure import storage as S

class AutoTests(unittest.TestCase):
    def setUp(self):
        self.tmp=tempfile.TemporaryDirectory(dir=os.environ.get('VZMK_TEST_PARENT'));self.parent=Path(self.tmp.name);self.root=self.parent/uuid4().hex;self.root.mkdir()
        for name in ('DEV','documents','validation'):(self.root/name).mkdir()
        d={'schema':'STAGE9_WORKSPACE_V1','workspace_id':self.root.name,'official_issue':False,'project':'Гараж Казахстан'};save(self.root/'STAGE9_WORKSPACE.json',d|{'seal':S.digest(d)})
        norms={0:[80],1:[150],2:[100,None],3:[100,80,60]}
        self.profiles={str(n):{'hashes':{f'ACT{n}':'HASH'},'layers':[{'ordinal':i+1,'required_raw':'∑1+2=180' if norm is None else norm,'required_individual_microns':norm} for i,norm in enumerate(rows)]} for n,rows in norms.items()}
        self.cfg={'schema':A.SCHEMA,'project':'Гараж Казахстан','enabled':True,'mode':A.MODE,'official_issue':False,'template_bindings':{n:t['hashes'] for n,t in self.profiles.items()},'climate':{'air_c':[15,25],'surface_c':[18,28],'relative_humidity':[40,65],'formula':A.FORMULA},'thickness':{str(n):[list(x) for x in pairs] for n,pairs in A.DEFAULT_THICKNESS.items()}}
    def tearDown(self):self.tmp.cleanup()
    def bad(self,change):
        cfg=copy.deepcopy(self.cfg);change(cfg)
        with self.assertRaises(S.Stop):A.validate_settings(cfg,self.profiles)
    def test_01_magnus_reference_vectors(self):
        for t,h,expected in [(20,50,9.255174598981256),(25,60,16.693149006198954),(0,100,0)]:self.assertAlmostEqual(A.dew_point(t,h),expected,places=10)
    def test_02_saturation_and_domain_endpoints(self):
        for t in (-45,-10,0,20,60):self.assertAlmostEqual(A.dew_point(t,100),t,places=10)
    def test_03_invalid_humidity(self):
        for h in (0,-1,101,True,'50',float('nan')):
            with self.assertRaises(S.Stop):A.dew_point(20,h)
    def test_04_invalid_temperature(self):
        for t in (-45.1,60.1,float('inf'),False,'20'):
            with self.assertRaises(S.Stop):A.dew_point(t,50)
    def test_05_result_outside_water_domain(self):
        with self.assertRaises(S.Stop):A.dew_point(-45,50)
    def test_06_defaults_are_valid(self):self.assertIs(A.validate_settings(self.cfg,self.profiles),self.cfg)
    def test_07_inverted_range(self):self.bad(lambda c:c['climate'].update(air_c=[25,15]))
    def test_08_boolean_range(self):self.bad(lambda c:c['climate'].update(air_c=[False,25]))
    def test_09_nan_range(self):self.bad(lambda c:c['climate'].update(air_c=[15,float('nan')]))
    def test_10_bad_precision(self):self.bad(lambda c:c['climate'].update(air_c=[15.01,25]))
    def test_11_impossible_surface_margin(self):self.bad(lambda c:c['climate'].update(surface_c=[-20,-10]))
    def test_12_unapproved_formula(self):self.bad(lambda c:c['climate'].update(formula='OLD_APPROXIMATION'))
    def test_13_official_mode(self):self.bad(lambda c:c.update(official_issue=True))
    def test_14_wrong_project(self):self.bad(lambda c:c.update(project='OTHER'))
    def test_15_wrong_template_pins(self):self.bad(lambda c:c['template_bindings'].update({'2':{'ACT2':'OTHER'}}))
    def test_16_wrong_layer_count(self):self.bad(lambda c:c['thickness'].update({'1':[[150,180],[150,180]]}))
    def test_17_norm_cannot_be_weakened(self):self.bad(lambda c:c['thickness'].update({'0':[[79,100]]}))
    def test_18_cumulative_norm(self):self.bad(lambda c:c['thickness'].update({'2':[[100,120],[70,100]]}))
    def test_19_cumulative_is_not_individual_norm(self):
        self.cfg['thickness']['2']=[[110,120],[70,90]];A.validate_settings(self.cfg,self.profiles)
        self.assertIsNone(self.profiles['2']['layers'][1]['required_individual_microns'])
    def test_20_variant_not_layer_count(self):
        self.assertEqual([len(self.cfg['thickness'][str(n)]) for n in range(4)],[1,1,2,3])
    def test_21_daily_cache_same_date_same_values(self):
        folder=self.parent/'cache';folder.mkdir();a=A.daily_climate(folder,'2026-10-04',self.cfg['climate']);b=A.daily_climate(folder,'2026-10-04',self.cfg['climate']);self.assertEqual(a,b)
    def test_22_daily_cache_changed_ranges_stop_without_overwrite(self):
        folder=self.parent/'cache';folder.mkdir();A.daily_climate(folder,'2026-10-04',self.cfg['climate']);p=folder/'days/2026-10-04.json';before=p.read_bytes()
        other=self.cfg['climate']|{'air_c':[16,25]}
        with self.assertRaisesRegex(S.Stop,'DAY_SETTINGS_CONFLICT'):A.daily_climate(folder,'2026-10-04',other)
        self.assertEqual(p.read_bytes(),before)
    def test_23_daily_tamper_stop(self):
        folder=self.parent/'cache';folder.mkdir();A.daily_climate(folder,'2026-10-04',self.cfg['climate']);p=folder/'days/2026-10-04.json';d=json.loads(p.read_text(encoding='utf8'));d['values']['surface_c']=-99;p.write_text(json.dumps(d),encoding='utf8')
        with self.assertRaisesRegex(S.Stop,'DAY_RECORD_DAMAGED'):A.daily_climate(folder,'2026-10-04',self.cfg['climate'])
    def test_24_sampling_ranges_and_dew_margin(self):
        folder=self.parent/'cache';folder.mkdir()
        for i in range(15):
            r=A.daily_climate(folder,f'2026-10-{i+1:02d}',self.cfg['climate']);v=r['values']
            for k in ('air_c','surface_c','relative_humidity'):self.assertTrue(self.cfg['climate'][k][0]<=v[k]<=self.cfg['climate'][k][1])
            self.assertGreaterEqual(v['surface_c']+1e-10,v['dew_point_raw_c']+3)
            self.assertFalse(r['official_issue']);self.assertEqual(r['mode'],'TEST_SIMULATION')
    def test_25_fixed_ranges(self):
        folder=self.parent/'cache';folder.mkdir();cfg=self.cfg['climate']|{'air_c':[20,20],'surface_c':[25,25],'relative_humidity':[50,50]};v=A.daily_climate(folder,'2026-10-04',cfg)['values'];self.assertEqual((v['air_c'],v['surface_c'],v['relative_humidity']),(20,25,50))
    def test_26_settings_persist_outside_templates(self):
        with patch.object(A,'profiles',return_value=self.profiles):A.store_settings(self.root,self.cfg);loaded=A.read_settings(self.root)
        self.assertEqual(loaded,self.cfg);self.assertTrue(A.settings_path(self.root).is_relative_to(self.root/'preferences'))
    def test_27_settings_lock_rejects_overlap(self):
        with A.settings_lock(self.root):
            with self.assertRaisesRegex(S.Stop,'LOCKED'):
                with A.settings_lock(self.root):pass
    def test_28_changed_disabled_settings_invalidate_preview(self):
        run=self.root/'validation/s9_000000000001';run.mkdir();d={'root':str(self.root),'run':str(run),'official_issue':False,'fingerprints':{},'akz_settings_sha256':None};plan={'data':d,'seal':S.digest(d)}
        core.verify(plan)
        with patch.object(A,'profiles',return_value=self.profiles):A.store_settings(self.root,self.cfg|{'enabled':False})
        with self.assertRaisesRegex(S.Stop,'SETTINGS_CHANGED_AFTER_PREVIEW'):core.verify(plan)
    def test_29_official_generator_rejected_before_io(self):
        with self.assertRaisesRegex(S.Stop,'OFFICIAL_GENERATION_FORBIDDEN'):A.generate_groups(self.root,None,[0],'s9_000000000001',mode='OFFICIAL')
    def test_30_invalid_extra_settings_keys(self):self.bad(lambda c:c.update(bypass=True))
    def test_31_thickness_preserves_whole_micron_display(self):self.bad(lambda c:c['thickness'].update({'0':[[80.1,100]]}))
    def test_32_shape_detects_visual_change_and_preserves_cumulative_norm(self):
        from openpyxl import Workbook
        from openpyxl.styles import Font
        p=self.parent/'shape.xlsx';b=Workbook();s=b.active;s.title='4';s['H26']='∑1+2=180';b.save(p)
        original=A.shape(p);self.assertEqual(original['4']['coating']['H26'],'∑1+2=180')
        s['A1'].font=Font(bold=True);b.save(p);changed=A.shape(p);b.close()
        self.assertNotEqual(original['4']['fonts'],changed['4']['fonts'])

if __name__=='__main__':unittest.main(verbosity=2)
