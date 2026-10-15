"""Read-only Stage 8 source inventory; excludes only this stage's new files."""
import hashlib
import json
import os
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]


def snapshot():
    roots = [ROOT, ROOT.parent / 'Гараж казахстан', ROOT.parents[1] / 'Другие шаблоны']
    excluded = [ROOT / 'stage8_closure', ROOT / 'tests/stage8']
    result = {}
    for top in roots:
        if not top.exists():
            continue
        for folder, dirs, files in os.walk(top, followlinks=False):
            here = Path(folder)
            dirs[:] = [d for d in dirs if not (here/d).is_symlink()
                       and not (here/d).is_junction() and d not in ('__pycache__', '.codebase-index')
                       and not any((here/d).is_relative_to(e) for e in excluded)
                       and not (here == ROOT/'validation' and d.startswith('stage8_'))]
            for name in files:
                p = here/name
                if p.is_symlink() or (here == ROOT and name in (
                        'stage8.py', 'STAGE8_DEV.md', 'Запустить Этап 8 DEV.cmd')):
                    continue
                result[str(p.resolve())] = hashlib.sha256(p.read_bytes()).hexdigest()
    for name in ('ВЗМК.xlsx', 'УЗК.xlsm'):
        p = ROOT.parents[1]/name
        if p.is_file():
            result[str(p.resolve())] = hashlib.sha256(p.read_bytes()).hexdigest()
    return result


if __name__ == '__main__':
    import sys
    target = Path(sys.argv[1])
    target.mkdir(parents=True, exist_ok=False)
    data = snapshot()
    (target/'SOURCE_BEFORE.json').write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding='utf8')
    print(json.dumps({'path': str(target.resolve()), 'guarded_files': len(data)}))
