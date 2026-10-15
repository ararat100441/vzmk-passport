"""Read-only baseline capture; never imports or writes production."""
from pathlib import Path
import hashlib
import json
import shutil

ROOT = Path(__file__).resolve().parents[2]
OUT = Path(__file__).resolve().parent
REVIEW = ROOT / 'validation/stage4_independent_review_20261007_0119'

def digest(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()

def write(name, value):
    (OUT/name).write_text(json.dumps(value, ensure_ascii=False, indent=2), encoding='utf-8')

if __name__ == '__main__':
    reviewed = json.loads((REVIEW/'reviewed_hashes_final.json').read_text(encoding='utf-8-sig'))
    actual = {p: digest(p) for p in reviewed}
    assert actual == reviewed, [p for p in reviewed if actual[p] != reviewed[p]]
    workbook = ROOT/'DEV/Справочники/Профили.xlsx'
    assert digest(workbook) == '8ff3ae1019c97d1a6a73151cd7ddf9b99be42c268cbf57ed1b56ce6285e6f615'
    assert digest(ROOT/'DEV/ВЗМК_DEV.xlsx') == '8e56f68518c70cade9efb09084dd67cd5d72e07e3fdc813b8515030e9cb98e80'
    mutable = {str(ROOT/p) for p in ['profiles_stage4/model.py', 'profiles_stage4/storage.py',
        'profiles_stage4/workbook.mjs', 'STAGE4_README.md']}
    generation = [workbook, *(workbook.parent/p for p in ['profiles.catalog.json','profiles.user.json','profiles.build.json'])]
    mutable.update(map(str, generation))
    protected = {p: h for p,h in reviewed.items() if p not in mutable}
    protected.update({str(p): digest(p) for p in REVIEW.rglob('*') if p.is_file()})
    write('reviewed_baseline.json', actual)
    write('protected_before.json', protected)
    write('dev_generation_before.json', {str(p):digest(p) for p in generation})
    snapshot = OUT/'baseline_generation'
    snapshot.mkdir()
    for p in generation:
        shutil.copy2(p, snapshot/p.name)
    print(json.dumps({'baseline':'PASS','reviewed_files':len(actual),'protected_files':len(protected),
                      'workbook_sha256':digest(workbook)}, ensure_ascii=False))
