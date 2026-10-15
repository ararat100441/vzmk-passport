import os,sys,json,hashlib,time,subprocess
from pathlib import Path
from uuid import uuid4
OUT=Path(__file__).resolve().parent;ROOT=OUT.parents[1];PRIOR=ROOT/'validation/akz_auto_20261013_02'
sys.path.insert(0,str(ROOT))
PY=r'C:\Users\arara\AppData\Local\Programs\Python\Python312\python.exe'
EXE=ROOT/'stage9_dist_0910/VZMK-Passport-0.9.10/VZMK-Passport-0.9.10.exe'
EXPECTED='7a53e4b9e388d4dcd66bd59cc4328ca4d3b74f69705de6dd69dc19489a602910'
def read(p):return json.loads(Path(p).read_text(encoding='utf-8-sig'))
def sha(p):
 h=hashlib.sha256()
 with Path(p).open('rb') as f:
  for b in iter(lambda:f.read(1048576),b''):h.update(b)
 return h.hexdigest()
def digest(x):return hashlib.sha256(json.dumps(x,ensure_ascii=False,sort_keys=True,separators=(',',':')).encode()).hexdigest()
def save(p,d):
 with Path(p).open('x',encoding='utf8') as f:json.dump(d,f,ensure_ascii=False,indent=2,default=str);f.flush();os.fsync(f.fileno())
def config():return read(OUT/'SETUP.json')
def environment():
 s=config();e=os.environ.copy();e.update(PATH=os.environ['SystemRoot']+'\\System32;'+os.environ['SystemRoot'],LOCALAPPDATA=s['base'],TEMP=s['temp'],TMP=s['temp'],PYTHONUTF8='1');e.pop('PYTHONHOME',None);e.pop('PYTHONPATH',None);return e
def task(label,name,expected='PASS',**kwargs):
 assert not (OUT/'FOREIGN_OFFICE_STOP.json').exists(),'BASELINE_OFFICE_EXIT_STOP'
 assert sha(EXE)==EXPECTED
 p=OUT/(label+'_'+uuid4().hex[:8]+'.json');cmd=[str(EXE),'--task',name,'--result',str(p)]
 for k,v in kwargs.items():
  cmd+=['--'+k.replace('_','-')]
  if v is not True:cmd+=[str(v)]
 save(p.with_suffix('.command.json'),{'cmd':cmd,'expected':expected,'exe_sha256':EXPECTED,'PATH':environment()['PATH'],'started':time.time()})
 proc=subprocess.run(cmd,env=environment(),capture_output=True,timeout=480)
 save(p.with_suffix('.process.json'),{'returncode':proc.returncode,'stdout':proc.stdout.decode('utf8','replace'),'stderr':proc.stderr.decode('utf8','replace'),'ended':time.time()})
 d=read(p);sup=read(d['office_supervisor']);assert d['status']==expected and (proc.returncode==0)==(expected=='PASS'),d
 assert sup['owned_orphans']==0 and not sup['ownership_errors'],sup
 assert all(r['exit_code']==0 and not r.get('forced_termination_attempted') for r in sup['instances']),sup
 assert not (OUT/'FOREIGN_OFFICE_STOP.json').exists(),'BASELINE_OFFICE_EXIT_STOP'
 print(label,name,expected,'Office',len(sup['instances']),flush=True)
 return d,p
