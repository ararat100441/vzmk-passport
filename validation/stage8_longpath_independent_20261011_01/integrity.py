import hashlib,json,sys,concurrent.futures,difflib,time
from pathlib import Path
RUN=Path(__file__).resolve().parent;ROOT=RUN.parents[1]
FIX=ROOT/'validation/stage8_longpath_fix_20261011_001'
def read(p):return json.loads(Path(p).read_text(encoding='utf-8-sig'))
def sha(p):
    h=hashlib.sha256()
    with Path(p).open('rb') as f:
        for b in iter(lambda:f.read(1024*1024),b''):h.update(b)
    return h.hexdigest()
def save(name,q):(RUN/name).write_text(json.dumps(q,ensure_ascii=False,indent=2),encoding='utf8')
def hashes(paths):
    def one(p):
        try:return p,sha(p)
        except Exception as e:return p,{'error':str(e)}
    out={}
    with concurrent.futures.ThreadPoolExecutor(max_workers=8) as ex:
        for i,(p,h) in enumerate(ex.map(one,sorted(paths))):
            out[p]=h
            if i%5000==0:print('HASHED',i,flush=True)
    return out
if __name__=='__main__':
    mode=sys.argv[1]
    if mode=='before':
        m=read(FIX/'REVIEW_MANIFEST.json');errors=[];count=0
        for k,v in m.items():
            if isinstance(v,dict):
                for p,h in v.items():
                    if isinstance(h,str) and len(h)==64 and Path(p).is_absolute():
                        count+=1
                        if sha(p)!=h:errors.append(p)
        product=m['product'];save('PRODUCT_BEFORE.json',hashes(product))
        assert not errors,errors
        changed=[];snap=[]
        for p,h in product.items():
            target=Path(p);after=FIX/'after_product'/target.name;before=FIX/'before_product'/target.name
            if after.exists():
                assert sha(after)==h,(p,'after snapshot mismatch')
                snap.append(str(after))
                if not before.exists() or sha(before)!=h:changed.append({'file':p,'before':sha(before) if before.exists() else 'NEW','after':h})
        save('PRODUCT_BINDING.json',{'manifest_sha256':sha(FIX/'REVIEW_MANIFEST.json'),'checked_manifest_entries':count,'errors':errors,'fingerprint':hashlib.sha256(json.dumps(product,ensure_ascii=False,sort_keys=True,separators=(',',':')).encode()).hexdigest(),'changed':changed,'snapshots':snap,'exact_diff_sha256':sha(FIX/'EXACT_PRODUCT.diff')})
        paths=set(read(FIX/'PROTECTED_BEFORE.json'))
        paths.update(str(p.resolve()) for p in ROOT.rglob('*') if p.is_file() and not p.is_relative_to(RUN) and '__pycache__' not in p.parts and '.git' not in p.parts)
        h=hashes(paths);save('PROTECTED_BEFORE.json',h)
        old=read(FIX/'PROTECTED_BEFORE.json');oldafter=read(FIX/'PROTECTED_AFTER.json')
        save('HISTORICAL_CLAIM.json',{'claimed_scope':len(old),'author_before_equals_after':old==oldafter,'current_match':sum(h[p]==v for p,v in old.items()),'current_drift':[p for p,v in old.items() if h[p]!=v]})
        print('BEFORE',len(h),'manifest',count,'changed',len(changed),flush=True)
    else:
        b=read(RUN/'PROTECTED_BEFORE.json');a=hashes(b);save('PROTECTED_AFTER.json',a)
        delta=[p for p,h in b.items() if a[p]!=h]
        volatile=[p for p in delta if '\\background-worker.' in p and '.lease\\heartbeat.' in p and '\\.codebase-index\\' in p]
        product=hashes(read(RUN/'PRODUCT_BEFORE.json'));save('PRODUCT_AFTER.json',product)
        save('SOURCE_INTEGRITY.json',{'checked':len(b),'changed':delta,'volatile_index_heartbeat':volatile,'protected_data_changed':[p for p in delta if p not in volatile],'read_errors':[p for p,v in a.items() if not isinstance(v,str)],'product_before_equals_after':product==read(RUN/'PRODUCT_BEFORE.json')})
        print('AFTER',len(a),'delta',delta,flush=True)
