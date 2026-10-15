"""Separate-process adversary; accepts paths only inside this review's fixtures."""
from pathlib import Path
import json
import os
import sys

base = Path(__file__).resolve().parent / 'fixtures'
mode, folder = sys.argv[1], Path(sys.argv[2]).resolve()
assert folder.is_relative_to(base.resolve())
names = ['Профили.xlsx', 'profiles.user.json', 'profiles.catalog.json', 'profiles.build.json']
results = []
if mode == 'vacancy':
    name = sys.argv[3]
    assert name in names
    with (folder/name).open('xb') as f:
        f.write(b'INDEPENDENT_EXTERNAL_VACANT_NAME')
    results.append({'file':name, 'created':True})
else:
    for name in names:
        for op in ['write','delete','rename','replace']:
            p = folder/name
            tmp = folder/(name+'.adversary')
            tmp.write_bytes(b'INDEPENDENT_EXTERNAL_REPLACEMENT')
            try:
                if op == 'write': p.write_bytes(b'INDEPENDENT_EXTERNAL_WRITE')
                elif op == 'delete': p.unlink()
                elif op == 'rename': p.rename(folder/(name+'.renamed'))
                else: os.replace(tmp,p)
            except OSError as exc:
                results.append({'file':name,'operation':op,'blocked':True,
                                'winerror':exc.winerror,'errno':exc.errno})
            else:
                results.append({'file':name,'operation':op,'blocked':False})
            finally:
                tmp.unlink(missing_ok=True)
sys.stdout.reconfigure(encoding='utf-8')
print(json.dumps(results,ensure_ascii=False))
