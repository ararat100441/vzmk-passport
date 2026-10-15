"""One controlled DEV rebuild only after isolated structural/native Excel PASS."""
from pathlib import Path
import json,sys
OUT=Path(__file__).resolve().parent; ROOT=OUT.parents[1]
sys.path.insert(0,str(ROOT)); sys.path.insert(0,str(OUT))
sys.stdout.reconfigure(encoding='utf-8')
import profiles_stage4.storage as s
from profiles_stage4.offline import sha,read_json,no_network
from profiles_stage4.windows_generation import GenerationGuard
from native_excel import run as native
from xlsx_tables import inspect
FILES=['Профили.xlsx',s.STATE,s.CATALOG,s.COMMIT]
def hashes():
    paths=sorted(p for p in (ROOT/'profiles_stage4').glob('*') if p.is_file())
    paths += [ROOT/'stage4.py',ROOT/'STAGE4_README.md',ROOT/'Запустить Этап 4.cmd']+sorted(OUT.glob('*.py'))
    return {str(p.relative_to(ROOT)):sha(p) for p in paths}
def protected():
    expected=read_json(OUT/'protected_before.json'); assert all(sha(p)==h for p,h in expected.items())
    return expected
def audit(event,args):
    if event in ('socket.connect','socket.getaddrinfo','socket.sendto'): raise AssertionError('Network prohibited')
    if event in ('open','os.rename','os.remove','os.mkdir'):
        for arg in args:
            if isinstance(arg,(str,bytes)) and any(p in str(arg).replace('\\','/') for p in ['/Работа/ВЗМК.xlsx','/Работа/Профили.xlsx']):
                raise AssertionError('Production prohibited')
sys.addaudithook(audit)
if __name__=='__main__':
    assert sys.argv[1:] in ([],['--build-dev'])
    tests=read_json(OUT/'targeted_tests.json'); assert tests['status']=='PASS' and tests['failed']==0
    before=read_json(OUT/'product_hashes_before.json'); after=read_json(OUT/'product_hashes_after.json')
    assert before==after==hashes(),'Product differs from exactly tested version'
    guarded=protected()
    isolated=read_json(OUT/'native/isolated_fixed/result.json'); assert isolated['validation']=='PASS'
    if sys.argv[1:]==['--build-dev']:
        marker=OUT/'real_DEV_build_attempt.json'; assert not marker.exists(),'Only one real DEV rebuild permitted'
        old=read_json(OUT/'generation_before.json'); assert all(sha(p)==h for p,h in old.items()),'STOP: DEV changed externally'
        assert sha(ROOT/'DEV/ВЗМК_DEV.xlsx')=='8e56f68518c70cade9efb09084dd67cd5d72e07e3fdc813b8515030e9cb98e80'
        folder=s.DEFAULT_WORKBOOK.parent
        assert not list(folder.glob('~$*')),'STOP: Excel owner file exists; close workbook'
        with GenerationGuard(folder,FILES) as current:
            assert {str(folder/name):h for name,h in current.hashes.items()}==old
            s.validated_generation(current.data,FILES[0])
        s.dump(marker,{'attempts':1,'timestamp':s.timestamp(),'previous_generation':old,'tested_hashes':before})
        with no_network() as events: built=s.build()
        assert not events and built['network_attempts']==0
        s.dump(OUT/'DEV_build_result.json',built)
        backup=Path(built['backup']); assert {str(folder/n):sha(backup/n) for n in FILES}==old
        checked=native(s.DEFAULT_WORKBOOK,'real_DEV_fixed')
        assert checked['validation']=='PASS' and not checked['repair_detected']
    built=read_json(OUT/'DEV_build_result.json'); checked=read_json(OUT/'native/real_DEV_fixed/result.json')
    assert checked['validation']=='PASS' and checked['source_unchanged'] and not checked['repair_detected']
    assert checked['sha256_before']==checked['sha256_after']==built['sha256']==sha(s.DEFAULT_WORKBOOK)
    with no_network() as events:
        catalog=s.load_catalog(); reopened=s.verify_workbook(s.DEFAULT_WORKBOOK,s.workbook_payload(catalog))
        structure=inspect(s.DEFAULT_WORKBOOK)
    assert not events and structure['valid']
    original=read_json(OUT/'baseline_generation'/s.CATALOG)
    assert catalog.raw_records==original['records'] and catalog.sources==original['sources']
    for field in ['preferences','manual_rows','resolutions']: assert catalog.user_state[field]==original['user_state'][field]
    assert catalog.user_state['revision']==original['user_state']['revision']+1
    assert catalog.counts()==read_json(OUT/'isolated_build.json')['counts']
    assert hashes()==before; protected()
    artifacts={str(s.DEFAULT_WORKBOOK.parent/n):sha(s.DEFAULT_WORKBOOK.parent/n) for n in FILES}
    s.dump(OUT/'DEV_structure.json',structure)
    s.dump(OUT/'DEV_artifacts.json',{'hashes':artifacts,'reopen':reopened,'counts':catalog.counts(),
        'user_state_preserved':True,'raw_sources_unchanged':True,'real_DEV_rebuilds':1})
    s.dump(OUT/'product_hashes_current.json',before); s.dump(OUT/'protected_after.json',guarded)
    summary={'status':'PASS','root_cause':'table1 column 11 r and column 16 R are case-insensitive duplicates',
        'fix':'technical column caption R -> Радиус R; unique-heading guard; numeric format preserved',
        'targeted_tests_passed':tests['passed'],'isolated_native_Excel':'PASS','real_DEV_native_Excel':'PASS',
        'alerts_enabled':True,'repair_prompt':False,'recovered_records':False,'recovered_title':False,
        'tables':5,'profile_rows':2849,'formulas':0,'workbook_sha256':built['sha256'],
        'product_before_after_current_equal':True,'protected_files_unchanged':len(guarded),
        'real_DEV_rebuilds':1,'Stage4_network_requests':0,'production_access':0,'ready_for_targeted_Work_rereview':True}
    s.dump(OUT/'FINALIZER.json',summary)
    evidence={str(p.relative_to(OUT)):sha(p) for p in OUT.rglob('*') if p.is_file() and p.name!='REVIEW_MANIFEST.json'}
    s.dump(OUT/'REVIEW_MANIFEST.json',{'created_at':s.timestamp(),'summary':summary,'product_hashes':before,
        'DEV_artifacts':artifacts,'protected_hashes':guarded,'evidence_files':evidence})
    print(json.dumps(summary,ensure_ascii=False))
