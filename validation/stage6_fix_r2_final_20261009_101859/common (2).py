from pathlib import Path
import sys,json,hashlib,shutil,os
from unittest.mock import patch
E=Path(__file__).resolve().parent;R=E.parents[1];sys.path.insert(0,str(R))
from akz_stage6 import core as c
from akz_stage6.gui import Launcher
ACTNAME='ТЕСТ АКЗ №0.xlsx';APPNAME='ТЕСТ Приложение к АКЗ №0.xlsx'
def load(p):return json.loads(Path(p).read_text('utf-8-sig'))
def save(n,x):(E/n).write_text(json.dumps(x,ensure_ascii=False,indent=2,default=str),encoding='utf-8')
def sha(p):return hashlib.sha256(Path(p).read_bytes()).hexdigest()
AUDIT={'production_attempts':[],'network_attempts':[]}
def auditor(event,args):
 if event=='open' and isinstance(args[0],(str,bytes,os.PathLike)):
  p=Path(os.fsdecode(args[0]))
  if p.name.casefold() in ('взмк.xlsx','узк.xlsm'):AUDIT['production_attempts'].append(str(p));raise AssertionError('PRODUCTION DENIED')
 if event in ('socket.connect','socket.getaddrinfo'):AUDIT['network_attempts'].append(str(args));raise AssertionError('NETWORK DENIED')
sys.addaudithook(auditor)
def setup(label):
 S=E/label;F=S/'fixtures';O=S/'output';F.mkdir(parents=True);O.mkdir()
 source=R/'validation/stage6_targeted_rereview_r1_r2_20261009/sandbox/fixtures'
 for p in source.iterdir():
  if p.is_file():shutil.copyfile(p,F/p.name)
 good=R/'validation/stage6_fix_r1_r2_20261009_094707/NEW_TEST_9_2026'
 shutil.copytree(good,O/'OLD_GOOD')
 (O/'unrelated.xlsx').write_bytes(b'UNRELATED USER FILE FIXTURE DO NOT DELETE')
 return F,O

def isolated(F,O):
 act=F/'ACT_template.xlsx';app=F/'APPLICATION_template.xlsx'
 return patch.multiple(c,REGISTRY=F/'DEV_registry.xlsx',ACT=act,APPLICATION=app,FIXTURES=F,OUTPUT=O,PINS={act:sha(act),app:sha(app)})
def snapshot(directory):return {str(p):sha(p) for p in directory.rglob('*') if p.is_file()}
def stop(fn,code):
 try:fn()
 except Exception as exc:
  assert code in str(exc),(code,type(exc).__name__,str(exc))
  return {'type':type(exc).__name__,'code':getattr(exc,'code',code),'message':str(exc),'cause':repr(exc.__cause__),'cleanup':repr(getattr(exc,'cleanup_error',None))}
 raise AssertionError('EXPECTED STOP '+code)
