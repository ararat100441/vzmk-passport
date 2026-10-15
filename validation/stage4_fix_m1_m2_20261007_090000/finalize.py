"""Bind targeted evidence to current code; permit exactly one real DEV rebuild."""
from pathlib import Path
import json
import sys

OUT=Path(__file__).resolve().parent
ROOT=OUT.parents[1]
sys.path.insert(0,str(ROOT))
sys.path.insert(0,str(OUT))
sys.stdout.reconfigure(encoding='utf-8')
import profiles_stage4.storage as s
from profiles_stage4.offline import read_json,sha,no_network
from profiles_stage4.model import calculate_uzk

def product_hashes():
    paths=sorted(p for p in (ROOT/'profiles_stage4').glob('*') if p.is_file())
    paths += [ROOT/'stage4.py',ROOT/'STAGE4_README.md',ROOT/'Запустить Этап 4.cmd']
    paths += sorted(OUT.glob('*.py'))
    return {str(p.relative_to(ROOT)):sha(p) for p in paths}

def protected_hashes():
    expected=read_json(OUT/'protected_before.json')
    current={p:sha(p) for p in expected}
    assert current==expected,'Protected Stage 1–3 / source / old review evidence changed'
    return current

def require_test_binding():
    result=read_json(OUT/'targeted_tests.json')
    assert result['status']=='PASS' and result['failed']==0 and result['real_DEV_builds']==0
    before=read_json(OUT/'product_hashes_before.json')
    after=read_json(OUT/'product_hashes_after.json')
    assert before==after==product_hashes(),'Current product differs from exact tested product'
    return result,before

def audit(event,args):
    if event in ('socket.connect','socket.getaddrinfo','socket.sendto'):
        raise AssertionError('Network prohibited')
    if event in ('open','os.rename','os.remove','os.mkdir'):
        for arg in args:
            if isinstance(arg,(str,bytes)):
                p=str(arg).replace('\\','/')
                if '/Работа/ВЗМК.xlsx' in p or '/Работа/Профили.xlsx' in p:
                    raise AssertionError('Production access prohibited')
sys.addaudithook(audit)

if __name__=='__main__':
    assert sys.argv[1:] in ([],['--build-dev'])
    result,tested=require_test_binding()
    protected=protected_hashes()
    if sys.argv[1:]==['--build-dev']:
        attempt=OUT/'real_dev_build_attempt.json'
        assert not attempt.exists(),'Exactly one controlled real DEV build is permitted'
        baseline=read_json(OUT/'dev_generation_before.json')
        assert baseline=={p:sha(p) for p in baseline},'Real DEV changed externally before controlled rebuild'
        assert sha(ROOT/'DEV/ВЗМК_DEV.xlsx')=='8e56f68518c70cade9efb09084dd67cd5d72e07e3fdc813b8515030e9cb98e80'
        s.load_catalog()  # Current four-file generation already hash-bound to baseline.
        s.dump(attempt,{'attempts':1,'timestamp':s.timestamp(),'previous_generation':baseline,'tested_hashes':tested})
        with no_network() as events:
            built=s.build()
        assert not events and built['network_attempts']==0
        backup=Path(built['backup'])
        assert {str(s.DEFAULT_WORKBOOK.parent/name):sha(backup/name) for name in
                [s.DEFAULT_WORKBOOK.name,s.STATE,s.CATALOG,s.COMMIT]}==baseline,'Backup is not the reviewed coherent generation'
        s.dump(OUT/'dev_build_result.json',built)
    built=read_json(OUT/'dev_build_result.json')
    with no_network() as events:
        catalog=s.load_catalog()
        reopened=s.verify_workbook(s.DEFAULT_WORKBOOK,s.workbook_payload(catalog))
        assert catalog.counts()==read_json(OUT/'post_fix_counts.json')
        for nd,want in [('ГОСТ Р 57837-2017',[[328,6],[348,9]]),('ГОСТ 26020-83',[[329,6.2],[310,8.5]])]:
            found=catalog.lookup('35Б1',nd)
            assert found['status']=='FOUND'
            assert [[r['L'],r['S']] for r in calculate_uzk(found)['rows']]==want
        assert catalog.lookup('35Б1')['status']=='FOUND_PREFERRED'
        assert catalog.lookup('20Ш1')['status']=='AMBIGUOUS'
    assert not events
    original=read_json(OUT/'baseline_generation'/s.STATE)
    assert catalog.user_state['revision']==original['revision']+1
    for field in ['manual_rows','preferences','resolutions']:
        assert catalog.user_state[field]==original[field]
    artifacts={str(s.DEFAULT_WORKBOOK.parent/name):sha(s.DEFAULT_WORKBOOK.parent/name) for name in
               [s.DEFAULT_WORKBOOK.name,s.STATE,s.CATALOG,s.COMMIT]}
    assert artifacts[str(s.DEFAULT_WORKBOOK)]==built['sha256']
    assert product_hashes()==tested
    protected_hashes()
    s.dump(OUT/'protected_after.json',protected)
    s.dump(OUT/'product_hashes_current.json',tested)
    s.dump(OUT/'dev_artifacts.json',{'hashes':artifacts,'reopen':reopened,'counts':catalog.counts(),
        'real_DEV_builds':1,'user_state_preserved':True,'35Б1_preference_unchanged':True})
    summary={'status':'PASS','M1':'PASS','M2':'PASS','targeted_tests_passed':result['passed'],
        'product_before_after_current_equal':True,'protected_files_unchanged':len(protected),
        'workbook_sha256':built['sha256'],'reopen':reopened,'real_DEV_rebuilds':1,'network_requests':0,
        'production_access':0,'ready_for_targeted_independent_rereview':True}
    s.dump(OUT/'FINALIZER.json',summary)
    # LAST write: the manifest binds every new evidence file to the tested current product.
    evidence={str(p.relative_to(OUT)):sha(p) for p in OUT.rglob('*')
              if p.is_file() and p.name!='REVIEW_MANIFEST.json'}
    s.dump(OUT/'REVIEW_MANIFEST.json',{'created_at':s.timestamp(),'summary':summary,'product_hashes':tested,
        'DEV_artifacts':artifacts,'protected_hashes':protected,'evidence_files':evidence})
    print(json.dumps(summary,ensure_ascii=False))
