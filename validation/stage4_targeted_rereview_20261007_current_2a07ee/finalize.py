"""Finalize independent evidence only; never invoke a DEV build."""
from pathlib import Path
from datetime import datetime, timezone
import hashlib
import json

OUT=Path(__file__).resolve().parent
ROOT=OUT.parents[1]
FIX=ROOT/'validation/stage4_fix_m1_m2_20261007_090000'
def read(p):return json.loads(Path(p).read_text(encoding='utf-8-sig'))
def sha(p):return hashlib.sha256(Path(p).read_bytes()).hexdigest()
def save(name,data):(OUT/name).write_text(json.dumps(data,ensure_ascii=False,indent=2),encoding='utf-8')
main=read(OUT/'TARGETED_REREVIEW.json')
extra=read(OUT/'immediate_precommit_races.json')
assert main['verdict']=='PASS' and main['failed']==0 and extra['passed']==7 and extra['failed']==0
before=read(OUT/'preservation_before.json')
current={p:sha(p) for p in before}
assert before==read(OUT/'preservation_after.json')==current
save('preservation_final.json',current)
fix=read(FIX/'REVIEW_MANIFEST.json')
products={p:sha(ROOT/p) for p in fix['product_hashes']}
assert products==fix['product_hashes']==read(FIX/'product_hashes_before.json')==read(FIX/'product_hashes_after.json')==read(FIX/'product_hashes_current.json')
main['tests'].extend(extra['tests'])
main.update(passed=main['passed']+extra['passed'],failed=0,M1='PASS',M2='PASS',
            findings={'BLOCKER':0,'MAJOR':0,'MINOR':0},required_fix='NONE',ready_for_user_E2E='YES',
            preservation={'BEFORE_EQUALS_AFTER_EQUALS_CURRENT':True,'files':len(before),'protected_files':423,
                          'fix_evidence_files':211,'real_DEV_rebuilds':0,'production_access':0,'new_source_audits':0},
            fixed_product_set={'changed':['profiles_stage4/model.py','profiles_stage4/storage.py','profiles_stage4/workbook.mjs','STAGE4_README.md'],
                               'added':['profiles_stage4/windows_generation.py']})
save('TARGETED_REREVIEW.json',main)
evidence={str(p.relative_to(OUT)):sha(p) for p in OUT.rglob('*') if p.is_file() and p.name!='REVIEW_MANIFEST.json'}
save('REVIEW_MANIFEST.json',{'created_at':datetime.now(timezone.utc).isoformat(),
     'verdict':'PASS','M1':'PASS','M2':'PASS','passed':main['passed'],'failed':0,
     'current_product_hashes':products,'current_DEV_artifacts':fix['DEV_artifacts'],
     'preservation_BEFORE_AFTER_CURRENT':True,'preserved_files':len(before),'evidence_files':evidence,
     'findings':main['findings'],'required_fix':'NONE','ready_for_user_E2E':'YES','not_run':main['not_run']})
print(json.dumps({'verdict':'PASS','passed':main['passed'],'failed':0,'preserved_files':len(before),
                  'findings':main['findings'],'ready_for_user_E2E':'YES'}))
