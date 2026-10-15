import os,sys,json,hashlib,subprocess,time
from pathlib import Path
from uuid import uuid4
OUT=Path(__file__).resolve().parent;ROOT=OUT.parents[1];sys.path.insert(0,str(ROOT))
def read(p):return json.loads(Path(p).read_text(encoding='utf8'))
def sha(p):return hashlib.sha256(Path(p).read_bytes()).hexdigest()
def save(p,d):
    with Path(p).open('x',encoding='utf8') as f:json.dump(d,f,ensure_ascii=False,indent=2,default=str);f.flush();os.fsync(f.fileno())
SETUP=read(OUT/'SETUP.json');WORKSPACE=Path(SETUP['workspace']);ENV=os.environ.copy();ENV['PATH']=os.environ['SystemRoot']+'\\System32;'+os.environ['SystemRoot'];ENV['LOCALAPPDATA']=SETUP['base'];ENV['TEMP']=ENV['TMP']=SETUP['temp'];ENV.pop('PYTHONPATH',None);ENV.pop('PYTHONHOME',None);ENV['PYTHONUTF8']='1'
PY=Path(r'C:\Users\arara\AppData\Local\Programs\Python\Python312\python.exe');EXE=ROOT/'stage9_dist_0910/VZMK-Passport-0.9.10/VZMK-Passport-0.9.10.exe'
def task(phase,name,expected='PASS',**args):
    assert not (OUT/'FOREIGN_OFFICE_STOP.json').exists(),'UNEXPECTED_FOREIGN_OFFICE_EXIT_STOP'
    prefix=[str(EXE)] if phase.startswith('EXE') else [str(PY),'-B',str(ROOT/'stage9.py')]
    if phase.startswith('EXE'):assert sha(EXE)==read(OUT/'BUILD_MANIFEST.json')['exe_sha256'],'EXE_SHA_CHANGED'
    token=uuid4().hex[:8];result=OUT/(phase+'_'+name+'_'+token+'.json');cmd=prefix+['--task',name,'--workspace',str(WORKSPACE),'--result',str(result)]
    for key,value in args.items():
        cmd+=['--'+key.replace('_','-')]
        if value is not True:cmd+=[str(value)]
    save(result.with_suffix('.command.json'),{'command':cmd,'runtime':'FROZEN_EXE' if phase.startswith('EXE') else 'PYTHON_SOURCE','exe_sha256':sha(EXE) if phase.startswith('EXE') else None,'PATH':ENV['PATH'],'started':time.time(),'expected':expected})
    p=subprocess.run(cmd,env=ENV,capture_output=True,timeout=450)
    save(result.with_suffix('.process.json'),{'returncode':p.returncode,'stdout':p.stdout.decode('utf8','replace'),'stderr':p.stderr.decode('utf8','replace'),'ended':time.time()})
    data=read(result);sup=read(data['office_supervisor']);assert sup['owned_orphans']==0 and not sup['ownership_errors'],sup
    assert all(not x.get('forced_termination_attempted') and x['exit_code']==0 for x in sup['instances']),sup
    assert data['status']==expected and (p.returncode==0)==(expected=='PASS'),data
    assert not (OUT/'FOREIGN_OFFICE_STOP.json').exists(),'UNEXPECTED_FOREIGN_OFFICE_EXIT_STOP'
    print(phase,name,data['status'],'Office',len(sup['instances']),'orphan 0',flush=True)
    return data,result,sup
