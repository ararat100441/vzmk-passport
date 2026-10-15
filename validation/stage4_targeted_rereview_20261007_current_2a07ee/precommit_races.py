"""Successful external mutations immediately before commit guard acquisition."""
from pathlib import Path
from unittest.mock import patch
import json
import hashlib
import shutil
import subprocess
import sys

OUT=Path(__file__).resolve().parent
ROOT=OUT.parents[1]
sys.path.insert(0,str(ROOT))
sys.dont_write_bytecode=True
sys.stdout.reconfigure(encoding='utf-8')
import profiles_stage4.storage as s
from profiles_stage4.windows_generation import GenerationGuard
def sha(p):return hashlib.sha256(Path(p).read_bytes()).hexdigest()
def read(p):return json.loads(Path(p).read_text(encoding='utf-8-sig'))
names=[s.DEFAULT_WORKBOOK.name,s.STATE,s.CATALOG,s.COMMIT]
def reuse(payload,dest,previews=None):
    shutil.copyfile(s.DEFAULT_WORKBOOK,dest)
    return {'network_attempts':0}
writer='''from pathlib import Path
import sys,json,zipfile,os
p=Path(sys.argv[1]).resolve();base=Path(sys.argv[3]).resolve()
assert p.is_relative_to(base)
op=sys.argv[2]
if op=='write':
 if p.suffix=='.xlsx':
  with zipfile.ZipFile(p,'a') as z:z.comment=b'EXTERNAL_AT_PRECOMMIT_GUARD'
 else:
  x=json.loads(p.read_text(encoding='utf-8'));x['precommit_external_marker']=True;p.write_text(json.dumps(x,ensure_ascii=False),encoding='utf-8')
elif op=='delete':p.unlink()
elif op=='rename':p.rename(p.with_name(p.name+'.external-renamed'))
elif op=='replace':
 q=p.with_name(p.name+'.external-replacement');q.write_bytes(p.read_bytes());
 with zipfile.ZipFile(q,'a') as z:z.comment=b'EXTERNAL_REPLACE_AT_PRECOMMIT_GUARD'
 os.replace(q,p)
print('external mutation successful')
'''
tests=[]
for i,(name,op) in enumerate([(n,'write') for n in names]+[(names[0],o) for o in ['delete','rename','replace']]):
    folder=OUT/'fixtures'/('immediate_precommit_'+str(i))/'DEV/Справочники'
    folder.mkdir(parents=True)
    for n in names:shutil.copyfile(s.DEFAULT_WORKBOOK.parent/n,folder/n)
    old={n:sha(folder/n) for n in names}
    original=GenerationGuard.__enter__
    count=[0];external={}
    def enter(guard):
        count[0]+=1
        if count[0]==2:
            result=subprocess.run([sys.executable,'-I','-B','-c',writer,str(folder/name),op,str(OUT/'fixtures')],capture_output=True,text=True,encoding='utf-8',timeout=20)
            assert result.returncode==0,result.stderr
            external.update({n:sha(folder/n) for n in names if (folder/n).exists()})
        return original(guard)
    with patch.object(s,'render_workbook',reuse),patch.object(GenerationGuard,'__enter__',enter):
        try:s.build(folder/names[0])
        except (OSError,ValueError) as exc:message=str(exc)
        else:raise AssertionError('Commit did not STOP')
    assert count[0]==2 and external
    assert {n:sha(folder/n) for n in names if (folder/n).exists()}==external
    assert all(external[n]==old[n] for n in names if n!=name)
    if op=='rename':assert sha(folder/(name+'.external-renamed'))==old[name]
    assert not (folder/'backups').exists()
    assert not list(folder.glob('.profiles-stage-*')) and not (folder/'.profiles.lock').exists()
    tests.append({'file':name,'operation':op,'status':'PASS','independent_process':True,
                  'phase':'immediately before second GenerationGuard acquisition, after render/reopen',
                  'STOP':message,'external_state_preserved':True,'partial_generation_absent':True,'backup_absent':True})
    print('PASS precommit '+name+' '+op,flush=True)
(OUT/'immediate_precommit_races.json').write_text(json.dumps({'passed':len(tests),'failed':0,'tests':tests},ensure_ascii=False,indent=2),encoding='utf-8')
