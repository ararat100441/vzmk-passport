"""Read-only bindings for the bounded Work 08 fix, before any product edits."""
import hashlib,json,shutil
from pathlib import Path
ROOT=Path(__file__).resolve().parents[3]
RUN=ROOT/'validation/stage8_fix_20261010_01'

def sha(p):return hashlib.sha256(Path(p).read_bytes()).hexdigest()
def read(p):return json.loads(Path(p).read_text(encoding='utf-8-sig'))
def save(p,v):Path(p).write_text(json.dumps(v,ensure_ascii=False,indent=2),encoding='utf8')

if __name__=='__main__':
    RUN.mkdir(exist_ok=False);(RUN/'before_product').mkdir()
    canonical=read(ROOT/'validation/stage8_20261010_03/REVIEW_MANIFEST.json')
    review=read(ROOT/'validation/stage8_independent_20261010_01/REVIEW_MANIFEST.json')
    print('independent manifest keys',list(review),flush=True)
    for p,h in canonical['product'].items():assert sha(p)==h,p
    for p,h in canonical['evidence'].items():assert sha(p)==h,p
    product={str(p):sha(p) for p in [*sorted((ROOT/'stage8_closure').glob('*.*')),ROOT/'stage8.py',ROOT/'STAGE8_DEV.md',ROOT/'Запустить Этап 8 DEV.cmd'] if p.is_file()}
    for p in product:
        q=Path(p);shutil.copyfile(q,RUN/'before_product'/q.name)
    save(RUN/'PRODUCT_CHECKPOINT.json',product)
    # Freeze existing source paths plus all previous Stage 8/Work08 evidence.
    source=read(ROOT/'validation/stage8_20261010_01/SOURCE_BEFORE.json')
    protected=set(source)
    for folder in (ROOT/'validation',ROOT/'tests/stage8'):
        for p in folder.rglob('*'):
            if p.is_file() and '__pycache__' not in p.parts and not p.is_relative_to(RUN) and 'work08_fix' not in p.parts:
                protected.add(str(p.resolve()))
    protected.difference_update(product)
    current={p:sha(p) for p in sorted(protected)}
    for p,h in source.items():assert current[p]==h,p
    save(RUN/'PROTECTED_BEFORE.json',current)
    save(RUN/'REVIEW_BINDINGS.json',{'canonical_manifest':sha(ROOT/'validation/stage8_20261010_03/REVIEW_MANIFEST.json'),
        'independent_manifest':sha(ROOT/'validation/stage8_independent_20261010_01/REVIEW_MANIFEST.json'),
        'independent_report':sha(ROOT/'validation/stage8_independent_20261010_01/FINAL_REPORT.md'),
        'canonical_product_match':True,'canonical_evidence_match':True,'protected_paths':len(current)})
    print('CHECKPOINT_PASS',len(current),len(product),flush=True)
