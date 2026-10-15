import json,hashlib,os,sys
from pathlib import Path
from datetime import datetime,timezone
from concurrent.futures import ThreadPoolExecutor
ROOT=Path(__file__).resolve().parents[2];OUT=Path(__file__).resolve().parent
def read(p):return json.loads(Path(p).read_text(encoding='utf-8-sig'))
def sha(p):
    h=hashlib.sha256()
    with Path(p).open('rb') as f:
        for b in iter(lambda:f.read(1024*1024),b''):h.update(b)
    return h.hexdigest()
manifest=OUT/'EVIDENCE_MANIFEST.json';assert not manifest.exists()
files={p.relative_to(OUT).as_posix():sha(p) for p in OUT.rglob('*') if p.is_file()}
old=ROOT/'validation/stage9_091_20261012_01';handoff=read(old/'EVIDENCE_MANIFEST.json')
external={str(old/p):h for p,h in handoff['files'].items()}|handoff['external_files']
for p in [old/'EVIDENCE_MANIFEST.json',old/'STAGE9_REPORT.md',old/'WORK09_ASSIGNMENT.md']:external[str(p)]=sha(p)
audit=read(OUT/'FINAL_PROCESS_AND_SOURCE_AUDIT_COMPLETE.json')
for base in [*(Path(p) for p in audit['data_roots']),*(Path(r['path']).parent for r in audit['reports'])]:
    for p in base.rglob('*'):
        if p.is_file():external[str(p)]=sha(p)
build=read(OUT/'BUILD_SHA256_AFTER.json');source=read(OUT/'SOURCE_SHA256_AFTER.json')
bindings={str(ROOT/p):h for p,h in source['files'].items()}
bindings.update({str(Path(build['exe']).parent/p):h for p,h in build['files'].items()})
immutable_bindings=dict(bindings)
protected=read(OUT/'PROTECTED_BEFORE.json');conflicts=[]
for p,expected in protected.items():
    actual=sha(p) if Path(p).is_file() else None
    if actual!=expected:conflicts.append({'path':p,'expected_sha256':expected,'actual_sha256':actual,'writer':'UNKNOWN'})
    if actual is not None:bindings[p]=actual
assert conflicts==[{k:read(OUT/'FINAL_INTEGRITY_CONFLICT.json')[k] for k in ('path','expected_sha256','actual_sha256','writer')}],conflicts
assert all(sha(p)==h for p,h in immutable_bindings.items()),'IMMUTABLE_PRODUCT_CHANGED'
refs={str(OUT/p):h for p,h in files.items()}|external|bindings
def verify(item):
    p,h=item;actual=sha(p) if Path(p).is_file() else None
    return {'path':p,'expected':h,'actual':actual} if actual!=h else None
with ThreadPoolExecutor(max_workers=8) as pool:bad=[v for v in pool.map(verify,refs.items()) if v]
assert not bad,bad
payload={'files':files,'external_files':external,'product_bindings':bindings}
aggregate=hashlib.sha256(json.dumps(payload,ensure_ascii=False,sort_keys=True,separators=(',',':')).encode('utf8')).hexdigest()
data={'schema':'WORK09_INDEPENDENT_EVIDENCE_V1','created_at_utc':datetime.now(timezone.utc).isoformat(),'review_root':str(OUT),'product_root':str(ROOT),'exe_sha256':build['exe_sha256'],'build_aggregate_sha256':build['aggregate_sha256'],'source_aggregate_sha256':source['aggregate_sha256'],'review_result':read(OUT/'REVIEW_RESULT.json'),'files':files,'external_files':external,'product_bindings':bindings,'protected_expectations':protected,'protected_integrity_conflicts':conflicts,'protected_integrity_status':'FAIL','immutable_product_integrity':'PASS','aggregate_sha256':aggregate,'manifest_written_last':True,'verified_observed_references':len(refs),'observed_reference_mismatches':0,'existing_55876_file_map':{'path':str(OUT/'INDEPENDENT_BEFORE.json'),'sha256':files['INDEPENDENT_BEFORE.json']},'after_integrity':{'path':str(OUT/'INDEPENDENT_AFTER_FINAL.json'),'sha256':files['INDEPENDENT_AFTER_FINAL.json']}}
with manifest.open('x',encoding='utf8') as f:json.dump(data,f,ensure_ascii=False,indent=2);f.flush();os.fsync(f.fileno())
loaded=read(manifest);assert loaded==data
print(json.dumps({'manifest':str(manifest),'sha256':sha(manifest),'review_files':len(files),'external_files':len(external),'product_bindings':len(bindings),'verified_references':len(refs),'mismatches':0,'manifest_written_last':True},ensure_ascii=False,indent=2),flush=True)
