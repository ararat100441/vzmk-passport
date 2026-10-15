"""Create the complete manifest last and verify it without further file writes."""
import concurrent.futures,hashlib,json,os,time
from pathlib import Path
OUT=Path(__file__).resolve().parent;ROOT=OUT.parents[1]
def read(p):return json.loads(Path(p).read_text(encoding='utf-8-sig'))
def sha(p):
    h=hashlib.sha256()
    with Path(p).open('rb') as f:
        for b in iter(lambda:f.read(1024*1024),b''):h.update(b)
    return h.hexdigest()
def digest(x):return hashlib.sha256(json.dumps(x,ensure_ascii=False,sort_keys=True,separators=(',',':')).encode('utf8')).hexdigest()
def save(p,x):
    with p.open('x',encoding='utf8') as f:json.dump(x,f,ensure_ascii=False,indent=2);f.flush();os.fsync(f.fileno())
assert not (OUT/'EVIDENCE_MANIFEST.json').exists()
setup=read(OUT/'SETUP.json'); protected=read(OUT/'PROTECTED_AFTER.json')['files']
paths={str(p) for r in (OUT,Path(setup['base']),Path(setup['parent'])) for p in r.rglob('*') if p.is_file()}
historic={}
for r in read(OUT/'HANDOFF_READBACK.json')['sets']:
    historic.update(read(r['manifest'])['files']);historic[r['manifest']]=r['sha256']
dev=read(OUT/'SOURCE_INTEGRITY_AFTER.json');devrefs={str(Path(dev['dev_source'])/p):h for p,h in dev['dev_files'].items()}
expected=protected|historic|devrefs
assert all(not (p in protected and protected[p]!=h) for p,h in historic.items())
paths.update(expected)
with concurrent.futures.ThreadPoolExecutor(10) as pool:files=dict(pool.map(lambda p:(p,sha(p)),sorted(paths)))
bad=[p for p,h in expected.items() if files[p]!=h];assert not bad,bad
controller=read(OUT/'CONTROLLER_SHA256_SEALED.json')['files'];assert all(files[p]==h for p,h in controller.items())
verification={'status':'PASS','mode':'full pre-seal hashes, manifest exclusive last, post-write readback+full rehash reported on stdout','files_before_seal_verify':len(files),'protected_files':len(protected),'historical_reference_files':len(historic),'sealed_dev_source_files':len(devrefs),'file_map_digest_before_seal_verify':digest(files),'mismatch':[],'EXE_sha256':read(OUT/'REVIEW_RESULT.json')['exe_sha256'],'product_review_verdict':'FAIL_SAFETY_STOP_NOT_RUN','no_files_written_after_manifest':True}
save(OUT/'SEAL_VERIFY.json',verification);files[str(OUT/'SEAL_VERIFY.json')]=sha(OUT/'SEAL_VERIFY.json')
payload={'schema':'WORK09_INDEPENDENT_EVIDENCE_MANIFEST_V1','review':'FINAL_INDEPENDENT_EXE0.9.9','created_unix':time.time(),'files':files,'file_count':len(files),'coverage':{'review_output':str(OUT),'external_owned_roots':[setup['base'],setup['parent']],'protected_ROOT_file_count':len(protected),'historical_self_evidence':'hash readback only; not independent runtime PASS','sealed_DEV_source':setup['source'],'production_registry_SHA_gate':False,'self':'excluded to avoid recursive SHA; canonical seal and file SHA returned by tool'},'review_result':str(OUT/'REVIEW_RESULT.json'),'final_verdict':'FAIL','reason':'Required runtime GUI and remaining fault cases NOT_RUN after safety STOP','BLOCKER':0,'MAJOR':0,'MINOR':0}
manifest=OUT/'EVIDENCE_MANIFEST.json';save(manifest,payload|{'seal':digest(payload)})
# No disk writes after this point, including no redirected stdout log.
m=read(manifest);assert m['seal']==digest({k:v for k,v in m.items() if k!='seal'})
with concurrent.futures.ThreadPoolExecutor(10) as pool:
    bad=[p for p,h,actual in pool.map(lambda item:(item[0],item[1],sha(item[0])),m['files'].items()) if h!=actual]
assert not bad,bad
current_owned={str(p) for r in (OUT,Path(setup['base']),Path(setup['parent'])) for p in r.rglob('*') if p.is_file()}
assert not current_owned-set(m['files'])-{str(manifest)},'UNSEALED_OWN_FILES'
print(json.dumps({'status':'PASS_MANIFEST_SEAL_AND_FULL_READBACK','manifest':str(manifest),'sha256':sha(manifest),'canonical_seal':m['seal'],'files':len(m['files']),'mismatch':bad,'last_file_written':str(manifest),'review_verdict':'FAIL','reason':'GUI/fault NOT_RUN safety STOP'},ensure_ascii=False,indent=2),flush=True)
