"""Independent process used ONLY against validation fixture paths."""
from pathlib import Path
import json
import os
import sys
sys.stdout.reconfigure(encoding='utf-8')

root = Path(__file__).resolve().parent
mode, folder = sys.argv[1], Path(sys.argv[2]).resolve()
assert folder.is_relative_to(root/'fixtures'), 'Fixture-only writer'
files = ['Профили.xlsx', 'profiles.user.json', 'profiles.catalog.json', 'profiles.build.json']
if mode == 'create_vacant':
    path = folder/files[0]
    with path.open('xb') as stream:
        stream.write(b'EXTERNAL_CREATION_AT_COMMIT_BOUNDARY')
    print(json.dumps({'created':True,'file':str(path)}))
else:
    results = []
    for name in files:
        for operation in ('write', 'delete', 'replace'):
            path = folder/name
            substitute = folder/(name+'.external-attempt')
            substitute.write_bytes(b'EXTERNAL_REPLACEMENT')
            try:
                if operation == 'write':
                    path.write_bytes(b'EXTERNAL_WRITE')
                elif operation == 'delete':
                    path.unlink()
                else:
                    os.replace(substitute, path)
            except OSError as exc:
                results.append({'file':name,'operation':operation,'blocked':True,'winerror':exc.winerror,'errno':exc.errno})
            else:
                results.append({'file':name,'operation':operation,'blocked':False})
            finally:
                substitute.unlink(missing_ok=True)
    print(json.dumps(results,ensure_ascii=False))
