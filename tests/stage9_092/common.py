import os,sys,subprocess
from pathlib import Path
from uuid import uuid4
ROOT=Path(__file__).resolve().parents[2];sys.path.insert(0,str(ROOT))
from stage9_app.workspace import save
from stage8_closure.storage import read,sha
OUT=ROOT/'validation/stage9_092_20261012_01'
SETUP=read(OUT/'SETUP.json');SOURCE=Path(SETUP['source']);VOM=Path(SETUP['vom'])
WORKSPACES=Path.home()/('V9D_092_'+Path(SETUP['base']).name.rsplit('_',1)[-1]);WORKSPACES.mkdir(exist_ok=True)
EXE=ROOT/'stage9_dist_097/VZMK-Passport-0.9.7/VZMK-Passport-0.9.7.exe'
ENV=os.environ.copy();ENV['PATH']=os.environ['SystemRoot']+'\\System32;'+os.environ['SystemRoot'];ENV.pop('PYTHONPATH',None);ENV.pop('PYTHONHOME',None)
ENV['LOCALAPPDATA']=SETUP['base'];TEMP=Path(SETUP['base'])/'new_temp';TEMP.mkdir(exist_ok=True);ENV['TEMP']=ENV['TMP']=str(TEMP)
def task(folder,name,expected='PASS',**args):
    path=folder/(name+'_'+uuid4().hex[:8]+'.json');command=[str(EXE),'--task',name,'--result',str(path)]
    for key,value in args.items():
        command.append('--'+key.replace('_','-'))
        if value is not True:command.append(str(value))
    proc=subprocess.run(command,env=ENV,timeout=360,capture_output=True)
    if not path.exists():raise AssertionError((name,proc.returncode,proc.stderr.decode('utf8','replace')))
    data=read(path);assert data['status']==expected and (proc.returncode==0)==(expected=='PASS'),data
    supervisor=read(data['office_supervisor']);assert supervisor['owned_orphans']==0 and not supervisor['ownership_errors'],supervisor
    if expected=='PASS':assert all(not r.get('forced_termination_attempted') and r['exit_code']==0 for r in supervisor['instances']),supervisor
    print(name,expected,'Office',len(supervisor['instances']),'orphans 0',flush=True)
    return data,supervisor,path
