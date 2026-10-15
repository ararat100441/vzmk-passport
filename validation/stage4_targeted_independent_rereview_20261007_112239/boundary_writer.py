from pathlib import Path
import json,os,sys
sys.stdout.reconfigure(encoding='utf-8')
folder=Path(sys.argv[1]).resolve()
base=Path(sys.argv[2]).resolve()
assert folder.is_relative_to(base/'fixtures')
results=[]
for name in ['Профили.xlsx','profiles.user.json','profiles.catalog.json','profiles.build.json']:
 for op in ['write','delete','rename','replace']:
  p=folder/name;q=folder/(name+'.independent-attempt')
  if op=='replace':q.write_bytes(b'EXTERNAL_REPLACEMENT')
  try:
   if op=='write':p.write_bytes(b'EXTERNAL_WRITE')
   elif op=='delete':p.unlink()
   elif op=='rename':p.rename(q)
   else:os.replace(q,p)
   results.append({'file':name,'operation':op,'blocked':False})
  except OSError as e:results.append({'file':name,'operation':op,'blocked':True,'winerror':e.winerror,'errno':e.errno})
  finally:q.unlink(missing_ok=True)
print(json.dumps(results,ensure_ascii=False))
