"""New independent evidence package. Only this new package is written."""
import json, hashlib, shutil, difflib
from pathlib import Path
RUN=Path(__file__).resolve().parent; ROOT=RUN.parents[1]
OLD=ROOT/'validation/stage8_recovery_independent_20261011_01'
FIX=ROOT/'validation/stage8_longpath_fix_20261011_001'
def read(p): return json.loads(p.read_text(encoding='utf-8-sig'))
def sha(p): return hashlib.sha256(p.read_bytes()).hexdigest()
def save(n,q): (RUN/n).write_text(json.dumps(q,ensure_ascii=False,indent=2),encoding='utf8')
bindings={}
for name in ('review.py','baseline_helpers.py','integrity.py','visible_gui.py'):
    source=OLD/name; text=source.read_text(encoding='utf8')
    if name=='review.py':
        text=text.replace("save(RUN/'NATIVE_EXACT.json'", "save(BATCH/'NATIVE_EXACT.json'").replace("save(RUN/'POSITIVE.json'", "save(BATCH/'POSITIVE.json'")
    if name=='integrity.py': text=text.replace('stage8_recovery_fix_20261011_01','stage8_longpath_fix_20261011_001')
    (RUN/name).write_text(text,encoding='utf8',newline='')
    bindings[name]={'source':str(source),'source_sha256':sha(source),'copy_sha256':sha(RUN/name),'adaptation':'Batch output paths only' if name=='review.py' else 'Author folder only' if name=='integrity.py' else 'None (newline normalization only)'}
save('HARNESS_BINDINGS.json',bindings)
before=read(OLD/'PRODUCT_AFTER.json'); after=read(FIX/'PRODUCT_CURRENT.json')
diff=[];changes=[]
for p,h in after.items():
    q=Path(p);a=FIX/'after_product'/q.name;b=FIX/'before_product'/q.name
    assert a.exists() and sha(a)==h
    if p in before: assert b.exists() and sha(b)==before[p],p
    else: assert not b.exists(),p
    assert sha(q)==h,p
    if before.get(p)!=h:
        changes.append({'file':p,'before':before.get(p,'NEW'),'after':h})
        diff.extend(difflib.unified_diff(b.read_text(encoding='utf8').splitlines(keepends=True) if b.exists() else [],a.read_text(encoding='utf8').splitlines(keepends=True),fromfile='before/'+q.name,tofile='after/'+q.name))
save('EXACT_DIFF_INDEPENDENT.json',{'changes':changes,'old_product_bound_to_prior_independent_checkpoint':True,'new_product_bound_to_current_files_and_author_snapshots':True,'provided_diff_sha256':sha(FIX/'EXACT_PRODUCT.diff')})
(RUN/'INDEPENDENT_PRODUCT.diff').write_text(''.join(diff),encoding='utf8',newline='')
canon=read(FIX/'CANONICAL_RESULTS.json')
assert canon['PASS']==80 and canon['FAIL']==0 and len(canon['cases'])==80
assert all(x['status']=='PASS' for x in canon['cases'])
save('AUTHOR_80_BINDING.json',{'canonical_sha256':sha(FIX/'CANONICAL_RESULTS.json'),'count':80,'status_counts':{'PASS':80,'FAIL':0},'manifest_product_fingerprint':read(FIX/'REVIEW_MANIFEST.json')['product_fingerprint'],'note':'Author claim; independent results are generated separately.'})
save('PREVIOUS_15_REQUIRED.json',[x['name'] for x in read(OLD/'RESULTS.json')['results'] if x['status']=='FAIL'])
print('BOOTSTRAP_READY',len(changes))
