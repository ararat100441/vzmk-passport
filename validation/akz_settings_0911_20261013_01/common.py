import os,sys,json,hashlib,subprocess,time,shutil
from pathlib import Path
from uuid import uuid4
OUT=Path(__file__).resolve().parent;ROOT=OUT.parents[1];sys.path.insert(0,str(ROOT))
def read(p):return json.loads(Path(p).read_text(encoding='utf8'))
def sha(p):return hashlib.sha256(Path(p).read_bytes()).hexdigest()
def save(p,d):
    with Path(p).open('x',encoding='utf8') as f:json.dump(d,f,ensure_ascii=False,indent=2,default=str);f.flush();os.fsync(f.fileno())
SETUP=read(OUT/'SETUP.json');PY=Path(r'C:\Users\arara\AppData\Local\Programs\Python\Python312\python.exe');EXE=ROOT/'stage9_dist_0911/VZMK-Passport-0.9.11/VZMK-Passport-0.9.11.exe'
ENV=os.environ.copy();ENV['PATH']=os.environ['SystemRoot']+'\\System32;'+os.environ['SystemRoot'];ENV.pop('PYTHONPATH',None);ENV.pop('PYTHONHOME',None);ENV['PYTHONUTF8']='1';ENV['LOCALAPPDATA']=SETUP['base'];ENV['TEMP']=ENV['TMP']=SETUP['temp']
def task(phase,name,expected='PASS',**args):
    assert not (OUT/'FOREIGN_OFFICE_STOP.json').exists(),'FOREIGN_OFFICE_EXIT_STOP'
    prefix=[SETUP['exe_0910']] if phase.startswith('OLD') else [str(EXE)] if phase.startswith('EXE') else [str(PY),'-B',str(ROOT/'stage9.py')]
    p=OUT/(phase+'_'+name+'_'+uuid4().hex[:8]+'.json');cmd=prefix+['--task',name,'--result',str(p)]
    for k,v in args.items():
        cmd+=['--'+k.replace('_','-')]
        if v is not True:cmd+=[str(v)]
    save(p.with_suffix('.command.json'),{'command':cmd,'exe_sha256':sha(prefix[0]),'runtime':'SOURCE' if prefix[0]==str(PY) else 'FROZEN_EXE','PATH':ENV['PATH'],'expected':expected,'started':time.time()})
    proc=subprocess.run(cmd,env=ENV,capture_output=True,timeout=600)
    save(p.with_suffix('.process.json'),{'returncode':proc.returncode,'stdout':proc.stdout.decode('utf8','replace'),'stderr':proc.stderr.decode('utf8','replace'),'ended':time.time()})
    d=read(p);s=read(d['office_supervisor']);assert s['owned_orphans']==0 and not s['ownership_errors'];assert all(not i.get('forced_termination_attempted') and i['exit_code']==0 for i in s['instances']),s
    assert d['status']==expected and (proc.returncode==0)==(expected=='PASS'),d
    assert not (OUT/'FOREIGN_OFFICE_STOP.json').exists(),'FOREIGN_OFFICE_EXIT_STOP'
    print(phase,name,d['status'],'Office',len(s['instances']),'orphan0',flush=True);return d,p
def bootstrap(label):
    d,p=task('OLD_SETUP_'+label,'import',source=SETUP['source'],parent=SETUP['parent']);work=Path(d['workspace'])
    source=Path(read(ROOT/'validation/akz_auto_20261013_02/SETUP.json')['workspace'])/'DEV/Stage6_Fixtures/assignments.xlsx';target=work/'DEV/Stage6_Fixtures/assignments.xlsx'
    before=sha(target);shutil.copyfile(source,target);assert sha(target)==sha(source)
    save(OUT/('WORKSPACE_'+label+'.json'),{'workspace':str(work),'fixture_copy_source':str(source),'fixture_sha256':sha(source),'initial_fixture_sha256':before,'existing_TEST_fixture_copied_without_content_edits':True})
    return work
