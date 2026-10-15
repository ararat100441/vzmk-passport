"""Independent read-only hash binding. Writes only this new review directory."""
import difflib, hashlib, json, sys
from pathlib import Path
RUN=Path(__file__).resolve().parent
ROOT=RUN.parents[1]
FIX=ROOT/'validation/stage8_fix_20261010_01'
def sha(p):
    h=hashlib.sha256()
    with Path(p).open('rb') as f:
        for b in iter(lambda:f.read(1024*1024),b''):h.update(b)
    return h.hexdigest()
def read(p):return json.loads(Path(p).read_text(encoding='utf-8-sig'))
def save(p,v):Path(p).write_text(json.dumps(v,ensure_ascii=False,indent=2),encoding='utf8')
def initial():
    manifest=read(FIX/'REVIEW_MANIFEST.json')
    verifications={}
    for section,values in manifest.items():
        if isinstance(values,dict) and values and all(isinstance(v,str) and len(v)==64 for v in values.values()):
            errors={}
            for p,h in values.items():
                try:
                    if sha(p)!=h:errors[p]='SHA_MISMATCH'
                except Exception as e:errors[p]=repr(e)
            verifications[section]={'checked':len(values),'errors':errors}
    prior=read(FIX/'PROTECTED_BEFORE.json')
    checked={};errors={}
    for p,h in prior.items():
        try:
            checked[p]=sha(p)
            if checked[p]!=h:errors[p]='SHA_MISMATCH'
        except Exception as e:errors[p]=repr(e)
    save(RUN/'CLAIMED_19618_CURRENT.json',{'checked':len(prior),'errors':errors,'hashes':checked})
    files=sorted(p for p in ROOT.rglob('*') if p.is_file() and not p.is_relative_to(RUN) and '__pycache__' not in p.parts and '.git' not in p.parts)
    snapshot={};read_errors={}
    for i,p in enumerate(files):
        try:snapshot[str(p.resolve())]=sha(p)
        except Exception as e:read_errors[str(p)]=repr(e)
        if i%5000==0:print('HASH',i,flush=True)
    save(RUN/'PROTECTED_BEFORE.json',snapshot)
    delivery=read(FIX/'FULL_PRODUCT_BINDING.json')['current']
    current={p:sha(p) for p in delivery}
    save(RUN/'PRODUCT_BEFORE.json',current)
    changes=[];diff=[]
    for p,h in read(FIX/'PRODUCT_CHECKPOINT.json').items():
        if sha(p)!=h:
            changes.append({'path':p,'before':h,'after':sha(p)})
            diff.extend(difflib.unified_diff((FIX/'before_product'/Path(p).name).read_text(encoding='utf8').splitlines(True),Path(p).read_text(encoding='utf8').splitlines(True),fromfile='before/'+Path(p).name,tofile='current/'+Path(p).name))
    (RUN/'INDEPENDENT_PRODUCT.diff').write_text(''.join(diff),encoding='utf8')
    save(RUN/'INITIAL_BINDING.json',{'manifest_sha256':sha(FIX/'REVIEW_MANIFEST.json'),'manifest_checks':verifications,'claimed_protected':len(prior),'claimed_mismatches':errors,'current_scope':len(snapshot),'read_errors':read_errors,'delivery_matches':current==delivery,'product_changes':changes})
    assert not errors and not read_errors and current==delivery
    assert all(not v['errors'] for v in verifications.values())
    print('INITIAL_PASS',len(prior),len(snapshot),len(changes),flush=True)
def merge():
    current=read(RUN/'PROTECTED_BEFORE.json');claimed=read(RUN/'CLAIMED_19618_CURRENT.json')['hashes']
    combined=claimed|current;save(RUN/'PROTECTED_BEFORE.json',combined)
    q=read(RUN/'INITIAL_BINDING.json');q['workspace_scope']=q['current_scope'];q['current_scope']=len(combined);q['claimed_scope_included_before_after']=True
    save(RUN/'INITIAL_BINDING.json',q);print('MERGED_SCOPE',len(combined),flush=True)
def final():
    from concurrent.futures import ThreadPoolExecutor
    before=read(RUN/'PROTECTED_BEFORE.json');after={};errors={}
    def check(item):
        p,h=item
        try:
            current=sha(p);return p,current,None if current==h else 'SHA_MISMATCH'
        except Exception as e:return p,None,repr(e)
    with ThreadPoolExecutor(max_workers=8) as pool:
        for i,(p,current,error) in enumerate(pool.map(check,before.items())):
            if current is not None:after[p]=current
            if error:errors[p]=error
            if i%5000==0:print('FINAL_HASH',i,flush=True)
    save(RUN/'PROTECTED_AFTER.json',after)
    product={p:sha(p) for p in read(RUN/'PRODUCT_BEFORE.json')}
    save(RUN/'PRODUCT_AFTER.json',product)
    save(RUN/'SOURCE_INTEGRITY.json',{'status':'PASS' if not errors else 'FAIL','protected_files':len(before),'changed_or_unreadable':errors,'current_product_matches':product==read(RUN/'PRODUCT_BEFORE.json'),'before_manifest_sha256':sha(RUN/'PROTECTED_BEFORE.json'),'after_manifest_sha256':sha(RUN/'PROTECTED_AFTER.json')})
    assert not errors and product==read(RUN/'PRODUCT_BEFORE.json')
    print('FINAL_INTEGRITY_PASS',len(before),flush=True)
if __name__=='__main__':{'before':initial,'merge':merge,'after':final}[sys.argv[1]]()
