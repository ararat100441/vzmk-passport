import os,sys,subprocess,json,time
from pathlib import Path
from uuid import uuid4
ROOT=Path(__file__).resolve().parents[2];OUT=Path(__file__).resolve().parent
state=OUT/'RUNTIME_ENV.json'
if not state.exists():
    local=Path(os.environ['USERPROFILE'])/'AppData/Local'/('VZMKWork09_'+uuid4().hex[:8]);local.mkdir()
    state.write_text(json.dumps({'LOCALAPPDATA':str(local),'controller_python':sys.executable,'runtime':'exact 0.9.1 EXE; minimal child PATH'},indent=2),encoding='utf8')
env=os.environ.copy();env.update(json.loads(state.read_text(encoding='utf8')))
env.update(STAGE9_VALIDATION_OUT=str(OUT),PYTHONDONTWRITEBYTECODE='1',PYTHONUTF8='1')
for name in sys.argv[1:]:
    tag='RECHECK_' if (OUT/(name+'_RUN.log')).exists() else ''
    logfile=OUT/(name+'_'+tag+'RUN.log')
    print('START',name,flush=True);start=time.monotonic()
    with logfile.open('x',encoding='utf8') as log:
        p=subprocess.run([sys.executable,'-B',str(ROOT/'tests/stage9_091'/(name+'.py'))],env=env,cwd=ROOT,stdout=log,stderr=subprocess.STDOUT)
    result={'script':name,'exit_code':p.returncode,'seconds':time.monotonic()-start,'log':str(logfile)}
    (OUT/(name+'_'+tag+'RUN.json')).write_text(json.dumps(result,indent=2),encoding='utf8')
    print(json.dumps(result),flush=True)
    print(logfile.read_text(encoding='utf8')[-4000:],flush=True)
    if p.returncode:sys.exit(p.returncode)
