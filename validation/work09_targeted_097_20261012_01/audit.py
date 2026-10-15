"""Independent read-only inventory and evidence checks; no product writes."""
from pathlib import Path
import sys, os, json, hashlib, concurrent.futures, datetime
OUT=Path(__file__).resolve().parent
ROOT=OUT.parents[1]
EXPECTED='cd140f337612c462903ecc35bb55173ba10839bb16fca61492c7ea7b65663cd7'
EXE=ROOT/'stage9_dist_097/VZMK-Passport-0.9.7/VZMK-Passport-0.9.7.exe'
def read(p): return json.loads(Path(p).read_text(encoding='utf-8'))
def save(p,obj):
    with Path(p).open('x',encoding='utf-8') as f: json.dump(obj,f,ensure_ascii=False,indent=2,default=str);f.flush();os.fsync(f.fileno())
def sha(p):
    h=hashlib.sha256()
    with Path(p).open('rb') as f:
        for b in iter(lambda:f.read(1024*1024),b''):h.update(b)
    return h.hexdigest()
def hashes(paths):
    with concurrent.futures.ThreadPoolExecutor(10) as pool:
        return dict(pool.map(lambda p:(str(p),sha(p)),paths))
def baseline():
    assert sha(EXE)==EXPECTED
    files=[p for p in ROOT.rglob('*') if p.is_file() and not p.is_relative_to(OUT) and '.codebase-index' not in p.parts]
    print('Protected files discovered',len(files),flush=True)
    data=hashes(files);save(OUT/'PROTECTED_BEFORE.json',{'files':data,'exe_sha256':sha(EXE),'production_registry_excluded':True})
    src=read(ROOT/'validation/stage9_092_20261012_01/DEV_SOURCE_COPY_INTEGRITY.json')
    actual={p:sha(Path(src['source'])/p) for p in src['files']};assert actual==src['files']
    save(OUT/'DEV_SOURCE_BEFORE.json',dict(src,independent_actual=actual))
    checks=[]
    for rel in ('validation/stage9_092_20261012_01/EVIDENCE_MANIFEST.json','validation/work09_independent_20261012_01/EVIDENCE_MANIFEST.json'):
        p=ROOT/rel;m=read(p);expected={}
        if 'external_files' in m:
            for field in ('files','external_files','product_bindings'):
                for name,value in m[field].items():
                    target=Path(name) if Path(name).is_absolute() else p.parent/name
                    expected[str(target)]=value
        else:expected=m['files']
        # Historical production observation is not an immutable gate.
        expected={k:v for k,v in expected.items() if Path(k).name!='ВЗМК.xlsx'}
        actual=hashes([Path(k) for k in expected]);bad=[k for k,v in expected.items() if actual[k]!=v]
        checks.append({'manifest':str(p),'manifest_sha256':sha(p),'references':len(expected),'mismatches':bad})
        assert not bad,checks[-1]
    save(OUT/'HANDOFF_READBACK.json',{'status':'PASS','checks':checks,'independent_runtime_proof':False})
    save(OUT/'BASELINE_READY.json',{'status':'PASS','files':len(data),'exe_sha256':sha(EXE),'utc':datetime.datetime.now(datetime.timezone.utc).isoformat()})
    print('BASELINE READY',len(data),flush=True)
if __name__=='__main__':baseline()
