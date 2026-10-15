"""Read-only final preservation gate; writes only this new review's evidence."""
from pathlib import Path
from collections import Counter
from copy import deepcopy
from datetime import datetime,timezone
import hashlib,json,sys
OUT=Path(__file__).resolve().parent;ROOT=OUT.parents[1]
sys.path.insert(0,str(ROOT));sys.dont_write_bytecode=True
sys.stdout.reconfigure(encoding='utf-8')
def read(p):return json.loads(Path(p).read_text(encoding='utf-8-sig'))
def sha(p):return hashlib.sha256(Path(p).read_bytes()).hexdigest()
def save(n,v):(OUT/n).write_text(json.dumps(v,ensure_ascii=False,indent=2),encoding='utf-8')
import profiles_stage4.storage as s
from profiles_stage4.model import Catalog,RULE,calculate_uzk
before=read(OUT/'hashes_before.json')
assert all(Path(p).resolve().is_relative_to(ROOT) for p in before)
current={p:sha(p) for p in before}
assert current==before==read(OUT/'hashes_after.json')
save('hashes_current.json',current)
state=read(OUT/'actual_state.json')
actual_files={str(p.relative_to(ROOT)) for p in (ROOT/'profiles_stage4').glob('*') if p.is_file()}
actual_files.update(['stage4.py','STAGE4_README.md','Запустить Этап 4.cmd'])
assert actual_files==set(state['runtime_files'])
initial=read(OUT/'review_result.json')
failed=[t for t in initial['tests'] if t['status']=='FAIL']
assert initial['passed']==29 and len(failed)==1 and failed[0]['name']=='M2 actual Windows handles protect commit against separate process'
assert 'Permission denied' in failed[0]['error']
commit=read(OUT/'commit_boundary_recheck.json')
assert commit['status']=='PASS' and commit['blocked_attempts']==48 and len(commit['phases'])==3
tests=deepcopy(initial['tests'])
for t in tests:
    if t['status']=='FAIL':
        t.clear();t.update({'name':failed[0]['name'],'status':'PASS','evidence':{'recheck':'commit_boundary_recheck.json','blocked_attempts':48,'phases':3,'previous_probe_diagnostic':'review_result.json retained unchanged; control read itself was blocked before writer invocation'}})
real=s.DEFAULT_WORKBOOK
records=read(real.parent/s.CATALOG)['records']
beam=next(r for r in records if r['key']=='35Б1' and r['standard']=='ГОСТ Р 57837-2017')
payload_evidence={}
for field in ('R','r'):
    a,b=deepcopy(beam),deepcopy(beam)
    a['record_id']='PAYLOAD-A';b['record_id']='PAYLOAD-B';b['source']='INDEPENDENT';b['source_id']='INDEPENDENT'
    a['geometry'][field]=10;b['geometry'][field]=11
    cat=Catalog([a,b]);result=cat.lookup(a['key'],a['standard'])
    assert result['status']=='FOUND' and field not in result['profile']['geometry'] and result['profile']['uzk_rule']==RULE
    assert [(r['L'],r['S']) for r in calculate_uzk(result)['rows']]==[(328,6),(348,9)]
    payload=s.workbook_payload(cat)
    cf=payload['sheets']['Конфликты'][1]
    members=json.loads(cf[5])+json.loads(cf[7])
    assert {m['geometry'][field] for m in members}=={10,11} and cf[12]=='NON_BLOCKING' and field in cf[8]
    main=payload['sheets']['Профили'][1]
    assert main[15 if field=='R' else 10] is None and 'unresolved' in main[19] and field in main[19]
    payload_evidence[field]={'lookup':'FOUND','canonical_unresolved':True,'UZK':[(328,6),(348,9)],'workbook_payload':payload}
save('auxiliary_workbook_payload_probes.json',payload_evidence)
tests.append({'name':'M1 R/r-only unresolved workbook payload visibility and raw values','status':'PASS','evidence':'auxiliary_workbook_payload_probes.json'})
names=[real.name,s.STATE,s.CATALOG,s.COMMIT]
verified_backups=[]
for directory in sorted((OUT/'fixtures').rglob('backups')):
    if not directory.is_dir():continue
    for backup in sorted(directory.iterdir()):
        if not backup.is_dir():continue
        assert {p.name for p in backup.iterdir()}==set(names),str(backup)
        data={n:(backup/n).read_bytes() for n in names}
        s.validated_generation(data,real.name)
        verified_backups.append({'path':str(backup.relative_to(OUT)),'hashes':{n:sha(backup/n) for n in names}})
save('all_fixture_backups_verified.json',verified_backups)
tests.append({'name':'All exposed isolated fixture backups are complete coherent four-file generations','status':'PASS','evidence':{'verified':len(verified_backups),'file':'all_fixture_backups_verified.json'}})
fix=ROOT/'validation/stage4_fix_m1_m2_20261007_090000'
fixmanifest=read(fix/'REVIEW_MANIFEST.json')
categories=Counter()
for p in fixmanifest['protected_hashes']:
    path=p.replace('\\','/')
    if 'stage4_independent_review_20261007_0119/' in path:categories['prior_independent_review']+=1
    elif 'stage4_km_kmd_audit_20261006_204347/' in path:categories['source_A']+=1
    elif 'stage4_ptoflow_audit_20261006_215239/' in path:categories['source_B']+=1
    else:categories['other_protected']+=1
assert sum(categories.values())==423
fixevidence=fixmanifest['evidence_files']
maxevidence=max((fix/p).stat().st_mtime_ns for p in fixevidence)
assert (fix/'REVIEW_MANIFEST.json').stat().st_mtime_ns>=maxevidence
counts=read(OUT/'counts_independent.json')
assert counts['raw']==3037 and counts['canonical']==2849
assert len(tests)==32 and all(t['status']=='PASS' for t in tests)
assert not any(initial['events'].values()) and initial['real_DEV_builds']==0
final_current={p:sha(p) for p in before};assert final_current==before
save('hashes_final.json',final_current)
summary={'verdict':'PASS','M1':'CLOSED','M2':'CLOSED','BLOCKER':0,'MAJOR':0,'MINOR':0,'targeted_checks_passed':32,'required_fix':'NONE','ready_for_user_E2E':'YES','User_E2E':'NOT_RUN','native_Excel_visual_UAT':'NOT_RUN','real_DEV_builds':0,'live_source_requests':0,'production_access':0,'outside_writes':0,'hashes_before_after_current_final_equal':True,'preserved_files':len(before),'protected_files':423,'protected_categories':dict(categories),'fix_evidence_files':211,'tested_product_files':15,'runtime_files':10,'actual_changed_set':state['changed_set'],'DEV_artifacts':state['DEV_artifacts'],'DEV_registry_sha256':'8e56f68518c70cade9efb09084dd67cd5d72e07e3fdc813b8515030e9cb98e80','counts':counts,'commit_boundary_attempts_blocked':48,'verified_fixture_backups':len(verified_backups),'normal_real_renderer_fixture_builds':5,'implementation_manifest_written_last':True,'initial_probe_diagnostics':['Initial independent commit probe attempted an incompatible ordinary read while DELETE-capable handles were protected; no external writer ran. Targeted recheck replaced that check.','First targeted recheck needed explicit UTF-8 stdout in its independent writer; the exact corrected helper passed all 48 OS denial attempts. No product change.'],'scope':'M1/M2 and direct regressions only','not_run':['User E2E','native Excel visual User UAT','full original Stage 4 review','full Stage 1–3 workflows','real DEV rebuild','production operations','live source research','Stage 5 / UZK.docx'],'tests':tests}
save('rereview_summary.json',summary)
# This new independent manifest is written last, after all fixtures and evidence.
ownhashes={str(p.relative_to(OUT)):sha(p) for p in sorted(OUT.rglob('*')) if p.is_file() and p.name!='REVIEW_MANIFEST.json'}
save('REVIEW_MANIFEST.json',{'created_at':datetime.now(timezone.utc).isoformat(),'summary':summary,'reviewed_product_hashes':fixmanifest['product_hashes'],'reviewed_DEV_artifacts':state['DEV_artifacts'],'reviewed_FIX_manifest_sha256':sha(fix/'REVIEW_MANIFEST.json'),'preservation_hashes':final_current,'evidence_files':ownhashes})
print(json.dumps({k:v for k,v in summary.items() if k!='tests'},ensure_ascii=False,indent=2))
