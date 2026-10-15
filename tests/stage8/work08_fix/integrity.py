"""Read-only protected-data hash proof, exact checkpoint diff and review bindings."""
import difflib,json,sys
from pathlib import Path
from stage8_closure import storage as S
RUN=S.ROOT/'validation/stage8_fix_20261010_01'
DEST=RUN/sys.argv[1] if len(sys.argv)>1 else RUN
if DEST!=RUN:DEST.mkdir(exist_ok=False)
before=S.read(RUN/'PROTECTED_BEFORE.json');current={};errors={}
for i,(p,h) in enumerate(before.items()):
    try:
        current[p]=S.sha(p)
        if current[p]!=h:errors[p]='SHA_CHANGED'
    except Exception as e:errors[p]=str(e)
    if i%4000==0:print('PROTECTED_HASH',i,flush=True)
S.write(DEST/'PROTECTED_AFTER.json',current)
canonical=S.read(S.ROOT/'validation/stage8_independent_20261010_01/REVIEW_MANIFEST.json')
checkpoint=S.read(RUN/'PRODUCT_CHECKPOINT.json')
reference=before|checkpoint
review_errors={}
for section in ('product','review_evidence'):
    for p,h in canonical[section].items():
        if reference.get(p)!=h:review_errors[p]='PRIOR_REVIEW_BINDING_MISMATCH'
changes=[];diff=[]
for p,old in checkpoint.items():
    new=S.sha(p)
    if new!=old:
        assert Path(p).name in ('core.py','storage.py','office.py','native.py','STAGE8_DEV.md'),p
        a=(RUN/'before_product'/Path(p).name).read_text(encoding='utf8').splitlines(True)
        b=Path(p).read_text(encoding='utf8').splitlines(True)
        diff.extend(difflib.unified_diff(a,b,fromfile='before/'+Path(p).name,tofile='after/'+Path(p).name))
        changes.append({'path':p,'before_sha256':old,'after_sha256':new})
(DEST/'EXACT_PRODUCT.diff').write_text(''.join(diff),encoding='utf8')
S.write(DEST/'CHANGED_PRODUCT_SHA.json',changes)
S.write(DEST/'SOURCE_INTEGRITY.json',{'status':'PASS' if not errors and not review_errors else 'FAIL',
    'protected_paths':len(before),'changed_protected':errors,'independent_review_bindings':review_errors,
    'source_before_sha':S.sha(RUN/'PROTECTED_BEFORE.json'),'source_after_sha':S.sha(DEST/'PROTECTED_AFTER.json'),
    'changed_product':len(changes)})
print('INTEGRITY',len(before),'CHANGES',len(errors),'REVIEW_BINDING_ERRORS',review_errors,flush=True)
assert not errors and not review_errors
