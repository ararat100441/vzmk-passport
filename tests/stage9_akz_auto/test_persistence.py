import copy,hashlib,json,sys,unittest
from pathlib import Path
from unittest.mock import patch
sys.path.insert(0,str(Path(__file__).resolve().parent))
import test_auto as fixtures
from stage9_app import akz_auto as A,akz_settings as K,core
from stage9_app.workspace import save
from stage8_closure import storage as S

class PersistenceTests(unittest.TestCase):
    def setUp(self):
        self.fixture=fixtures.AutoTests();self.fixture.setUp()
        for name in ('root','parent','cfg','profiles'):setattr(self,name,getattr(self.fixture,name))
    def tearDown(self):self.fixture.tearDown()
    def test_numeric_identity_and_real_change(self):
        a={'range':[15,25],'nested':[0,0.1]};b={'nested':[-0.0,0.10],'range':[15.00,25.0]}
        self.assertEqual(K.canonical_json(a),K.canonical_json(b));self.assertEqual(K.digest(a),K.digest(b))
        self.assertNotEqual(K.digest(a),K.digest(a|{'range':[16,26]}))
    def test_no_decimal_rounding_in_canonicalization(self):
        self.assertEqual(K.canonical_json([15.1,1e-7,100]),'[15.1,0.0000001,100]')
        self.assertNotEqual(K.digest([15.1]),K.digest([15.100000000000001]))
    def test_gui_comma_and_equivalent_notations(self):
        for text in ('15','15.0','15,00','1.5e1'):self.assertEqual(K.gui_number(text),15)
        self.assertEqual(K.gui_number('15,1'),15.1)
        for text in ('NaN','Infinity','15.00000000000000000001'):
            with self.assertRaises(S.Stop):K.gui_number(text)
    def test_equal_save_retains_exact_bytes_and_history(self):
        with patch.object(A,'profiles',return_value=self.profiles):
            A.store_settings(self.root,self.cfg);path=A.settings_path(self.root);before=path.read_bytes();history=list(path.parent.glob('history/*.json'))
            cfg=copy.deepcopy(self.cfg);cfg['climate']['air_c']=[15.0,25.0];A.store_settings(self.root,cfg)
            self.assertEqual(path.read_bytes(),before);self.assertEqual(list(path.parent.glob('history/*.json')),history)
    def test_old_float_day_is_reused_without_rewriting(self):
        folder=self.parent/'days_test';folder.mkdir();cfg=copy.deepcopy(self.cfg['climate']);cfg['air_c']=[15.0,25.0]
        old=A.daily_climate(folder,'2026-10-04',cfg);old.pop('settings_digest_schema');old.pop('climate_settings');old.pop('seal');old['settings_digest']=S.digest(cfg);old['seal']=S.digest(old)
        path=folder/'days/2026-10-04.json';path.write_text(json.dumps(old),encoding='utf8');before=path.read_bytes()
        self.assertEqual(A.daily_climate(folder,'2026-10-04',self.cfg['climate']),old);self.assertEqual(path.read_bytes(),before)
        with self.assertRaisesRegex(S.Stop,'SETTINGS_CONFLICT'):A.daily_climate(folder,'2026-10-04',cfg|{'air_c':[16,26]})
    def test_snapshot_bound_exact_sha_and_future_settings(self):
        run=self.root/'validation/s9_000000000001';run.mkdir()
        with patch.object(A,'profiles',return_value=self.profiles):
            A.store_settings(self.root,self.cfg);p=A.settings_path(self.root);ref=K.freeze(self.root,run,S.sha(p));expected=K.verify_snapshot(self.root,run,ref)
            changed=copy.deepcopy(self.cfg);changed['climate']['air_c']=[16,26];A.store_settings(self.root,changed)
            self.assertEqual(K.verify_snapshot(self.root,run,ref),expected)
            with self.assertRaisesRegex(S.Stop,'EXISTS_NO_OVERWRITE'):K.freeze(self.root,run,S.sha(p))
    def test_snapshot_tamper_safe_stop(self):
        import win32api,win32con
        run=self.root/'validation/s9_000000000001';run.mkdir()
        with patch.object(A,'profiles',return_value=self.profiles):
            A.store_settings(self.root,self.cfg);ref=K.freeze(self.root,run,S.sha(A.settings_path(self.root)));p=Path(ref['path'])
            win32api.SetFileAttributes(str(p),win32con.FILE_ATTRIBUTE_NORMAL);p.write_bytes(p.read_bytes()+b' ')
            with self.assertRaisesRegex(S.Stop,'SHA_CONFLICT'):K.verify_snapshot(self.root,run,ref)
    def test_snapshot_substitution_other_run_stops(self):
        run=self.root/'validation/s9_000000000001';run.mkdir();other=self.root/'validation/s9_000000000002';other.mkdir()
        with patch.object(A,'profiles',return_value=self.profiles):
            A.store_settings(self.root,self.cfg);ref=K.freeze(self.root,run,S.sha(A.settings_path(self.root)))
            with self.assertRaisesRegex(S.Stop,'SHA_CONFLICT'):K.verify_snapshot(self.root,other,ref)
    def test_legacy_recovery_is_exact_sha_and_read_only(self):
        p=self.root/'preferences/akz_auto/settings.json';p.parent.mkdir(parents=True);history=p.parent/'history';history.mkdir()
        original=json.dumps(self.cfg,ensure_ascii=False,indent=2).encode('utf8');expected=hashlib.sha256(original).hexdigest();p.write_text('{}',encoding='utf8')
        save(history/'original.json',{'settings':self.cfg});before={str(x):x.read_bytes() for x in p.parent.rglob('*.json')}
        proof=K.legacy_evidence(p,expected);self.assertEqual(proof['status'],'LEGACY_SETTINGS_RECOVERED_EXACT_SHA')
        self.assertEqual({str(x):x.read_bytes() for x in p.parent.rglob('*.json')},before)
        with self.assertRaisesRegex(S.Stop,'UNPROVEN'):K.legacy_evidence(p,'0'*64)
    def test_legacy_recovery_windows_writer_crlf_is_exact_and_read_only(self):
        p=self.root/'preferences/akz_auto/settings.json';p.parent.mkdir(parents=True);history=p.parent/'history';history.mkdir()
        save(p,self.cfg);original=p.read_bytes();self.assertIn(b'\r\n',original)
        expected=hashlib.sha256(original).hexdigest();save(history/'original.json',{'settings':self.cfg})
        p.write_bytes(b'{}');before={str(x):x.read_bytes() for x in p.parent.rglob('*.json')}
        proof=K.legacy_evidence(p,expected);self.assertEqual(proof['status'],'LEGACY_SETTINGS_RECOVERED_EXACT_SHA')
        self.assertEqual(proof['sha256'],expected)
        self.assertEqual({str(x):x.read_bytes() for x in p.parent.rglob('*.json')},before)
    def test_legacy_numeric_equivalence_is_not_exact_byte_evidence(self):
        p=self.root/'preferences/akz_auto/settings.json';p.parent.mkdir(parents=True);history=p.parent/'history';history.mkdir()
        changed=copy.deepcopy(self.cfg);changed['climate']['air_c']=[15.0,25.0]
        expected=hashlib.sha256(json.dumps(changed,ensure_ascii=False,indent=2).replace('\n','\r\n').encode('utf8')).hexdigest()
        save(history/'integer.json',{'settings':self.cfg});p.write_bytes(b'{}')
        with self.assertRaisesRegex(S.Stop,'UNPROVEN'):K.legacy_evidence(p,expected)
    def test_manifest_cannot_omit_snapshot_binding(self):
        run=self.root/'validation/s9_000000000001';run.mkdir()
        with patch.object(A,'profiles',return_value=self.profiles):
            A.store_settings(self.root,self.cfg);ref=K.freeze(self.root,run,S.sha(A.settings_path(self.root)));m={'inputs':{},'records':[],'akz_settings_snapshot':ref}
            with self.assertRaisesRegex(S.Stop,'NOT_BOUND'):core.manifest_settings_evidence(m,run/'MANIFEST.json')

if __name__=='__main__':unittest.main(verbosity=2)
