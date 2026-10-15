import json,hashlib,os,sys,time
from pathlib import Path
ROOT=Path(__file__).resolve().parents[2]
OUT=Path(__file__).resolve().parent
def sha(p):
    h=hashlib.sha256()
    with p.open('rb') as f:
        for b in iter(lambda:f.read(1024*1024),b''):h.update(b)
    return h.hexdigest()
def save(name,data):
    with (OUT/name).open('x',encoding='utf8') as f:json.dump(data,f,ensure_ascii=False,indent=2)
mode=sys.argv[1]
if mode=='before':
    paths=[p for p in ROOT.rglob('*') if p.is_file() and not p.is_relative_to(OUT) and '.codebase-index' not in p.parts]
    external=json.loads((OUT/'PROTECTED_BEFORE.json').read_text(encoding='utf8'))
    paths=sorted(set(paths)|{Path(p) for p in external})
    data={str(p):{'sha256':sha(p),'bytes':p.stat().st_size} for p in paths}
    save('INDEPENDENT_BEFORE.json',data)
    save('BASELINE_SUMMARY.json',{'files':len(data),'bytes':sum(v['bytes'] for v in data.values()),'root':str(ROOT),'external_pins':len(external)})
    print('BASELINE',len(data),'files',flush=True)
else:
    before=json.loads((OUT/'INDEPENDENT_BEFORE.json').read_text(encoding='utf8'))
    from concurrent.futures import ThreadPoolExecutor
    def check(item):
        p,v=item;actual=sha(Path(p)) if Path(p).is_file() else None
        return {'path':p,'before':v,'after':actual} if actual!=v['sha256'] else None
    with ThreadPoolExecutor(max_workers=8) as pool:bad=[v for v in pool.map(check,before.items()) if v]
    actual={str(p) for p in ROOT.rglob('*') if p.is_file() and not p.is_relative_to(OUT) and '.codebase-index' not in p.parts}
    save('INDEPENDENT_AFTER_FINAL.json' if mode=='after_final' else 'INDEPENDENT_AFTER.json',{'status':'FAIL' if bad else 'PASS','checked':len(before),'changed_or_missing':bad,'new_files_outside_review':sorted(actual-set(before))})
    print('AFTER',len(before),'changed',len(bad),'new',len(actual-set(before)),flush=True)
