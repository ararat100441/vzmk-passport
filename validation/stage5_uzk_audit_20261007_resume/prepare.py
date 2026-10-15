import hashlib, json, shutil
from pathlib import Path
from datetime import datetime, timezone, timedelta

OUT=Path(__file__).resolve().parent
APP=OUT.parent.parent
OLD=APP/'validation'/'stage5_uzk_audit_20261007_stop'
SOURCE=APP.parent.parent/'УЗК.xlsm'
DEV=APP/'DEV'/'УЗК_DEV.xlsm'
CHECKPOINT=APP/'DEV'/'DEV_UZK_CHECKPOINT.json'
def sha(p):
    h=hashlib.sha256()
    with Path(p).open('rb') as f:
        for b in iter(lambda:f.read(1048576),b''):h.update(b)
    return h.hexdigest()
def save(p,data):p.write_text(json.dumps(data,ensure_ascii=False,indent=2),encoding='utf-8')
assert SOURCE.is_file(), 'STOP: exact authoritative source absent'
oldmanifest=json.loads((OLD/'EVIDENCE_MANIFEST.json').read_text(encoding='utf-8-sig'))
oldchecks={n:sha(OLD/n)==v for n,v in oldmanifest['evidence_sha256'].items()}
assert all(oldchecks.values()), 'STOP: previous audit evidence mismatch'
baseline=json.loads((OLD/'accepted_baseline_check.json').read_text(encoding='utf-8'))
for row in baseline:
    assert sha(row['path'])==row['expected'], 'STOP: accepted baseline drift '+row['path']
source_before=sha(SOURCE)
created=False
if DEV.exists() or CHECKPOINT.exists():
    assert DEV.is_file() and CHECKPOINT.is_file(), 'STOP: partial existing DEV copy/checkpoint'
    checkpoint=json.loads(CHECKPOINT.read_text(encoding='utf-8-sig'))
    assert sha(DEV)==checkpoint['dev_initial_sha256'], 'STOP: existing DEV changed'
else:
    with SOURCE.open('rb') as src,DEV.open('xb') as dst:
        shutil.copyfileobj(src,dst)
    dev_sha=sha(DEV)
    source_after=sha(SOURCE)
    assert source_before==dev_sha==source_after, 'STOP: copy integrity mismatch'
    stat=SOURCE.stat()
    checkpoint={'source_path':str(SOURCE),'source_sha256':source_before,
                'dev_path':str(DEV),'dev_initial_sha256':dev_sha,
                'creation_timestamp':datetime.now(timezone(timedelta(hours=4))).isoformat(),
                'source_size_bytes':stat.st_size,
                'source_modified_timestamp':datetime.fromtimestamp(stat.st_mtime,timezone(timedelta(hours=4))).isoformat(),
                'source_sha256_after_copy':source_after,
                'policy':{'development_input':'DEV_ONLY','auto_sync':False,
                          'future_refresh':'EXPLICIT_USER_DECISION_ONLY','registry_decrement':'NOT_AUTHORIZED',
                          'source_use':'INITIAL_COPY_AND_FINGERPRINT_ONLY'}}
    with CHECKPOINT.open('x',encoding='utf-8') as f:json.dump(checkpoint,f,ensure_ascii=False,indent=2)
    created=True
protected={r['path']:sha(r['path']) for r in baseline}
oldprotected=json.loads((OLD/'preservation_before.json').read_text(encoding='utf-8'))
for p,h in oldprotected.items():
    item=Path(p)
    if item.is_relative_to(APP/'passport_app') or item.is_relative_to(APP/'profiles_stage4') or item.parent==APP:
        assert sha(item)==h, 'STOP: accepted code drift '+p
        protected[p]=h
protected.update({str(DEV):sha(DEV),str(CHECKPOINT):sha(CHECKPOINT),str(SOURCE):source_before})
save(OUT/'preservation_before.json',protected)
save(OUT/'prepare.json',{'created_dev_copy':created,'checkpoint':checkpoint,'old_evidence_checks':oldchecks,
                         'accepted_baseline_checks':len(baseline),'protected_files':len(protected)})
print(json.dumps({'created':created,'source_sha':source_before,'dev_sha':sha(DEV),'accepted_checks':len(baseline)},ensure_ascii=True))
