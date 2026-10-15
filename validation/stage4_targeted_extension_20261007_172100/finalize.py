"""One controlled real DEV publication, then a read-only integrity finalizer."""
from pathlib import Path
from decimal import Decimal
import argparse,json,sys
OUT=Path(__file__).resolve().parent; ROOT=OUT.parents[1]
sys.path.insert(0,str(ROOT)); sys.path.insert(0,str(OUT)); sys.stdout.reconfigure(encoding='utf-8')
from profiles_stage4 import storage as s
from profiles_stage4.model import calculate_uzk
from profiles_stage4.offline import read_json,sha
from profiles_stage4.windows_generation import GenerationGuard
from native_excel import run as native
from xlsx_tables import inspect
FILES=['Профили.xlsx',s.STATE,s.CATALOG,s.COMMIT]
def product_hashes():
    paths=sorted(p for p in (ROOT/'profiles_stage4').glob('*') if p.is_file())
    paths += [ROOT/'stage4.py',ROOT/'STAGE4_README.md',ROOT/'Запустить Этап 4.cmd']+sorted(OUT.glob('*.py'))
    return {str(p.relative_to(ROOT)):sha(p) for p in paths}
def preserve():
    expected=read_json(OUT/'protected_before.json')
    for p,h in expected.items(): assert sha(p)==h,p
    return expected
def main():
    parser=argparse.ArgumentParser(); parser.add_argument('--build-dev',action='store_true'); args=parser.parse_args()
    tests=read_json(OUT/'targeted_tests.json'); assert tests['status']=='PASS' and tests['failed']==0
    before=read_json(OUT/'product_hashes_before.json'); after=read_json(OUT/'product_hashes_after.json')
    assert before==after==product_hashes(), 'Current code differs from tested code'
    protected=preserve()
    if args.build_dev:
        assert not (OUT/'real_DEV_build_attempt.json').exists(), 'Only one real DEV build is authorized'
        expected=read_json(OUT/'generation_before.json')
        assert not list(s.DEFAULT_WORKBOOK.parent.glob('~$*')), 'Close workbook before DEV update'
        assert not (s.DEFAULT_WORKBOOK.parent/'.profiles.lock').exists()
        with GenerationGuard(s.DEFAULT_WORKBOOK.parent,FILES) as guard:
            assert {str(s.DEFAULT_WORKBOOK.parent/n):h for n,h in guard.hashes.items()}==expected
            previous=s.validated_generation(guard.data,s.DEFAULT_WORKBOOK.name)
        state=read_json(OUT/'proposed_user_state.json'); assert state['revision']==previous['revision']
        s.dump(OUT/'real_DEV_build_attempt.json',{'attempts':1,'timestamp':s.timestamp(),'baseline_hashes':expected,'product_hashes':before,'workbook_closed_preflight':True})
        result=s.build(user_state=state)
        assert result['network_attempts']==0
        backup=Path(result['backup'])
        assert {str(s.DEFAULT_WORKBOOK.parent/n):sha(backup/n) for n in FILES}==expected
        s.dump(OUT/'real_DEV_build.json',result)
        native_result=native(s.DEFAULT_WORKBOOK,'real_DEV')
        assert native_result['validation']=='PASS' and not native_result['repair_detected']
        s.dump(OUT/'real_DEV_native.json',native_result)
    assert (OUT/'real_DEV_build.json').is_file()
    artifacts={str(s.DEFAULT_WORKBOOK.parent/n):sha(s.DEFAULT_WORKBOOK.parent/n) for n in FILES}
    build=read_json(OUT/'real_DEV_build.json'); assert artifacts[str(s.DEFAULT_WORKBOOK)]==build['sha256']
    if (OUT/'DEV_artifacts.json').exists(): assert artifacts==read_json(OUT/'DEV_artifacts.json')
    native_result=read_json(OUT/'real_DEV_native.json')
    assert native_result['validation']=='PASS' and native_result['source_unchanged']
    assert artifacts[str(s.DEFAULT_WORKBOOK)]==native_result['sha256_after']==native_result['sha256_before']
    catalog=s.load_catalog(); reopen=s.verify_workbook(s.DEFAULT_WORKBOOK,s.workbook_payload(catalog))
    structure=inspect(s.DEFAULT_WORKBOOK); assert structure['valid']
    preview=[]
    for mark,raw,q,expected in [('Бп2-1','[16П',8,[[1280,5],[1024,8.4]]),('Пс1-9','Гн.[140X60X4',9,[[3600,4]]),('Пс1-10','Гн.[140X60X4',15,[[6000,4]]),('Пс2-16','Гн.[]120х3',1,[[480,3]])]:
        result=catalog.lookup(raw); assert result['status']=='FOUND'
        uzk=calculate_uzk(result); single=[[r['L'],r['S']] for r in uzk['rows']]
        total=[[float(Decimal(str(r['L']))*q),r['S']] for r in uzk['rows']]
        assert total==expected and all(r['frequency_mhz']==5 and r['angle_degrees']==70 for r in uzk['rows'])
        preview.append({'mark':mark,'raw_profile':raw,'quantity':q,'lookup_status':result['status'],'standard':result['profile']['standard'],'single':single,'total':total,'frequency_mhz':5,'angle_degrees':70})
    s.dump(OUT/'four_mark_preview.json',{'scope':'READ_ONLY_VALIDATION_PREVIEW; no Stage 5 product code','marks':preview})
    lookup_results={raw:{'lookup':catalog.lookup(raw),'UZK':calculate_uzk(catalog.lookup(raw))} for raw in ['[16П','Гн.[140X60X4','Гн.[]120х3']}
    s.dump(OUT/'required_lookup_results.json',lookup_results)
    s.dump(OUT/'DEV_artifacts.json',artifacts)
    current=product_hashes(); assert current==before==after
    s.dump(OUT/'product_hashes_current.json',current)
    summary={'status':'PASS','targeted_tests_passed':tests['passed'],'counts':catalog.counts(),'protected_files_unchanged':len(protected),
             'product_before_after_current_equal':True,'real_DEV_rebuilds':1,'native_Excel':'PASS','repair':False,'network_requests':0,'production_access':0,'reopen':reopen,'ready_for_targeted_independent_review':True}
    s.dump(OUT/'final_summary.json',summary)
    report='\n'.join(['# Stage 4 targeted extension — PASS','',
        'User decision: 2026-10-07; KMD facts explicitly supplied and confirmed by user.',
        'Two verified manual records; three persisted UZK rule scopes; safe lookup-only channel alias.',
        'All source A/B records, old geometry/status/provenance and 8 blocking / 7 non-blocking disagreements unchanged.',
        'Tests: '+str(tests['passed'])+' PASS; isolated mutations first; real DEV rebuilds: 1.',
        'Native Excel: normal Open, read-only, alerts enabled, 5 tables, no dialog/log/recovered title; file SHA unchanged.',
        'Counts: '+json.dumps(catalog.counts(),ensure_ascii=False),
        'Workbook SHA-256: '+artifacts[str(s.DEFAULT_WORKBOOK)],
        'BEFORE = AFTER = CURRENT for tested product and validation scripts.',
        'Protected files unchanged: '+str(len(protected)),
        'Four-mark preview: '+json.dumps(preview,ensure_ascii=False),
        'NOT_RUN: Stage 5 implementation, Word/UZK.docx, XLSM work, new source audits, web, production, full Stage 1–3 regression, GUI User E2E, KMD re-audit.',
        'See targeted_tests.json, user_confirmed_decision.json, required_lookup_results.json, real_DEV_build.json and native/real_DEV/result.json.',
    ])
    (OUT/'REVIEW_REPORT.md').write_text(report+'\n',encoding='utf-8')
    assert product_hashes()==before
    evidence={str(p.relative_to(OUT)):{'sha256':sha(p),'size':p.stat().st_size} for p in OUT.rglob('*') if p.is_file() and '__pycache__' not in p.parts and p.name!='REVIEW_MANIFEST.json'}
    s.dump(OUT/'REVIEW_MANIFEST.json',{'created_at':s.timestamp(),'summary':summary,'product_hashes':current,'DEV_artifacts':artifacts,'protected_hashes':protected,'evidence_files':evidence})
    print(json.dumps({'summary':summary,'workbook_sha256':artifacts[str(s.DEFAULT_WORKBOOK)],'manifest_sha256':sha(OUT/'REVIEW_MANIFEST.json')},ensure_ascii=False))
if __name__=='__main__': main()
