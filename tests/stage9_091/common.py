import os,sys,subprocess
from pathlib import Path
from uuid import uuid4
ROOT=Path(__file__).resolve().parents[2];sys.path.insert(0,str(ROOT))
from stage9_app.workspace import save
from stage8_closure.storage import read,sha
OUT=Path(os.environ.get('STAGE9_VALIDATION_OUT',str(ROOT/'validation/stage9_091_20261012_01')))
OUT.mkdir(parents=True,exist_ok=True)
EXE=ROOT/'stage9_dist_091/VZMK-Passport-0.9.1/VZMK-Passport-0.9.1.exe'
ENV=os.environ.copy();ENV['PATH']=os.environ['SystemRoot']+'\\System32;'+os.environ['SystemRoot'];ENV.pop('PYTHONPATH',None);ENV.pop('PYTHONHOME',None)
def task(folder,name,expected='PASS',**args):
    path=folder/(name+'_'+uuid4().hex[:8]+'.json');command=[str(EXE),'--task',name,'--result',str(path)]
    for key,value in args.items():
        command.append('--'+key.replace('_','-'))
        if value is not True:command.append(str(value))
    proc=subprocess.run(command,env=ENV,timeout=360,capture_output=True)
    if not path.exists():raise AssertionError((name,proc.returncode,proc.stderr.decode('utf8','replace')))
    data=read(path);assert data['status']==expected and (proc.returncode==0)==(expected=='PASS'),data
    supervisor=read(data['office_supervisor']);assert supervisor['owned_orphans']==0 and not supervisor['ownership_errors'],supervisor
    print(name,expected,'Office',len(supervisor['instances']),'orphans 0',flush=True)
    return data,supervisor,path
