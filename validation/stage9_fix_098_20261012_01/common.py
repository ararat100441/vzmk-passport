"""Independent frozen-task controller. Controllers never substitute source runtime."""
import os,sys,subprocess,json,hashlib,time
from pathlib import Path
from uuid import uuid4
OUT=Path(__file__).resolve().parent;ROOT=OUT.parents[1];sys.path.insert(0,str(ROOT))
def read(p):return json.loads(Path(p).read_text(encoding='utf8'))
def save(p,obj):
    with Path(p).open('x',encoding='utf8') as f:json.dump(obj,f,ensure_ascii=False,indent=2,default=str);f.flush();os.fsync(f.fileno())
def sha(p):return hashlib.sha256(Path(p).read_bytes()).hexdigest()
SETUP=read(OUT/'SETUP_099.json');SOURCE=Path(SETUP['source']);VOM=Path(SETUP['vom']);WORKSPACES=Path(SETUP['parent'])
EXE=ROOT/'stage9_dist_099/VZMK-Passport-0.9.9/VZMK-Passport-0.9.9.exe'
EXPECTED=read(OUT/'NEW_EXE_BINDING.json')['exe_sha256']
ENV=os.environ.copy();ENV['PATH']=os.environ['SystemRoot']+'\\System32;'+os.environ['SystemRoot'];ENV.pop('PYTHONPATH',None);ENV.pop('PYTHONHOME',None)
ENV['LOCALAPPDATA']=SETUP['base'];ENV['TEMP']=ENV['TMP']=SETUP['temp']
def task(folder,name,expected='PASS',**args):
    assert sha(EXE)==EXPECTED
    path=folder/(name+'_'+uuid4().hex[:8]+'.json');cmd=[str(EXE),'--task',name,'--result',str(path)]
    for key,value in args.items():
        cmd.append('--'+key.replace('_','-'))
        if value is not True:cmd.append(str(value))
    command=path.with_suffix('.command.json');save(command,{'command':cmd,'exe_sha256':EXPECTED,'expected_status':expected,'PATH':ENV['PATH'],'LOCALAPPDATA':ENV['LOCALAPPDATA'],'started':time.time()})
    proc=subprocess.run(cmd,env=ENV,timeout=450,capture_output=True)
    save(path.with_suffix('.process.json'),{'returncode':proc.returncode,'stdout':proc.stdout.decode('utf8','replace'),'stderr':proc.stderr.decode('utf8','replace'),'ended':time.time()})
    assert path.exists(),(name,proc.returncode,proc.stderr)
    data=read(path);assert data['status']==expected and (proc.returncode==0)==(expected=='PASS'),data
    sup=read(data['office_supervisor']);assert sup['owned_orphans']==0 and not sup['ownership_errors'],sup
    if expected=='PASS':assert all(not r.get('forced_termination_attempted') and r['exit_code']==0 for r in sup['instances']),sup
    print(name,expected,'Office',len(sup['instances']),'orphan 0',flush=True)
    return data,sup,path
