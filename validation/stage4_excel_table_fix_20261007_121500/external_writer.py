"""Bounded independent-process commit probe; validation fixtures ONLY."""
from pathlib import Path
import json,os,sys
sys.stdout.reconfigure(encoding='utf-8')
folder=Path(sys.argv[1]).resolve()
assert folder.is_relative_to(Path(__file__).resolve().parent/'fixtures')
results=[]
for name in ['Профили.xlsx','profiles.user.json','profiles.catalog.json','profiles.build.json']:
    for operation in ['write','delete','replace']:
        target=folder/name; substitute=folder/(name+'.external-probe')
        substitute.write_bytes(b'EXTERNAL_PROBE')
        try:
            if operation=='write': target.write_bytes(b'EXTERNAL_PROBE')
            elif operation=='delete': target.unlink()
            else: os.replace(substitute,target)
        except OSError as exc: results.append({'file':name,'operation':operation,'blocked':True,'winerror':exc.winerror,'errno':exc.errno})
        else: results.append({'file':name,'operation':operation,'blocked':False})
        finally: substitute.unlink(missing_ok=True)
print(json.dumps(results,ensure_ascii=False))
