"""Bind the accepted fix and independent failure before any recovery edits."""
import json, shutil
from pathlib import Path
from stage8_closure import storage as S

RUN=S.ROOT/'validation/stage8_recovery_fix_20261011_01'
if __name__=='__main__':
    RUN.mkdir(exist_ok=False);(RUN/'before_product').mkdir()
    prior=S.read(S.ROOT/'validation/stage8_fix_20261010_01/REVIEW_MANIFEST.json')
    review=S.read(S.ROOT/'validation/stage8_independent_targeted_20261011_01/REVIEW_MANIFEST.json')
    bindings={}
    for name,manifest in [('fix',prior),('review',review)]:
        errors=[];count=0
        for section,values in manifest.items():
            if isinstance(values,dict):
                for p,h in values.items():
                    if isinstance(h,str) and len(h)==64 and Path(p).is_absolute():
                        count+=1
                        if S.sha(p)!=h:errors.append(p)
        bindings[name]={'checked':count,'errors':errors}
        assert not errors,errors
    product={str(p):S.sha(p) for p in [*sorted((S.ROOT/'stage8_closure').glob('*.*')),S.ROOT/'stage8.py',S.ROOT/'STAGE8_DEV.md',S.ROOT/'Запустить Этап 8 DEV.cmd'] if p.is_file()}
    for p in product:shutil.copyfile(p,RUN/'before_product'/Path(p).name)
    S.write(RUN/'PRODUCT_BEFORE.json',product)
    protected=set(S.read(S.ROOT/'validation/stage8_20261010_01/SOURCE_BEFORE.json'))
    for base in [S.ROOT/'validation',S.ROOT/'tests/stage8']:
        protected.update(str(p.resolve()) for p in base.rglob('*') if p.is_file() and '__pycache__' not in p.parts and not p.is_relative_to(RUN) and 'recovery_fix' not in p.parts)
    protected.difference_update(product)
    hashes={p:S.sha(p) for p in sorted(protected)}
    S.write(RUN/'PROTECTED_BEFORE.json',hashes)
    S.write(RUN/'PRIOR_BINDINGS.json',bindings)
    print('CHECKPOINT',len(product),len(hashes),bindings,flush=True)
