"""Independent process; writes only inside this rereview's isolated fixtures."""
from pathlib import Path
import json, os, sys
base=Path(__file__).resolve().parent
folder=Path(sys.argv[1]).resolve()
assert folder.is_relative_to(base/'fixtures')
names=['Профили.xlsx','profiles.user.json','profiles.catalog.json','profiles.build.json']
if sys.argv[2]=='vacancy':
    with (folder/names[0]).open('xb') as f:
        f.write(b'INDEPENDENT_EXTERNAL_CREATION_MUST_SURVIVE')
    print(json.dumps({'created':True}))
else:
    results=[]
    for name in names:
        path=folder/name
        for op in ['write','delete','rename','replace']:
            dest=folder/(name+'.external-probe')
            if op=='replace': dest.write_bytes(b'EXTERNAL_REPLACEMENT')
            try:
                if op=='write': path.write_bytes(b'EXTERNAL_WRITE')
                elif op=='delete': path.unlink()
                elif op=='rename': path.rename(dest)
                else: os.replace(dest,path)
                results.append({'file':name,'operation':op,'blocked':False})
            except OSError as e:
                results.append({'file':name,'operation':op,'blocked':True,'winerror':e.winerror})
            finally:
                dest.unlink(missing_ok=True)
    print(json.dumps(results,ensure_ascii=False))
