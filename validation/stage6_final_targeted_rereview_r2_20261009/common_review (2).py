from pathlib import Path
import sys,os,json,hashlib,shutil,copy,ast,difflib,traceback
from unittest.mock import patch
from decimal import Decimal
sys.stdout.reconfigure(encoding='utf-8')
E=Path(__file__).resolve().parent;ROOT=E.parents[1]
FIX=ROOT/'validation/stage6_fix_r2_final_20261009_101859'
PREVIOUS=ROOT/'validation/stage6_targeted_rereview_r1_r2_20261009'
sys.path.insert(0,str(ROOT))
AUDIT={'production_attempts':[],'network_attempts':[]}
def auditor(event,args):
    if event=='open' and isinstance(args[0],(str,bytes,os.PathLike)):
        p=Path(os.fsdecode(args[0]))
        if p.name.casefold() in ('взмк.xlsx','узк.xlsm'):
            AUDIT['production_attempts'].append(str(p));raise AssertionError('PRODUCTION DENIED')
    if event in ('socket.connect','socket.getaddrinfo'):
        AUDIT['network_attempts'].append(str(args));raise AssertionError('NETWORK DENIED')
sys.addaudithook(auditor)
def load(p):return json.loads(Path(p).read_text('utf-8-sig'))
def save(n,v):(E/n).write_text(json.dumps(v,ensure_ascii=False,indent=2,default=str),encoding='utf-8')
def sha(p):return hashlib.sha256(Path(p).read_bytes()).hexdigest()
def snapshot(p):return {str(f):sha(f) for f in Path(p).rglob('*') if f.is_file()}
ACTNAME='ТЕСТ АКЗ №0.xlsx';APPNAME='ТЕСТ Приложение к АКЗ №0.xlsx'
PAIR=ROOT/'validation/stage6_fix_r1_r2_20261009_094707/NEW_TEST_9_2026'
EXPECTED={ACTNAME:'b640796570a8d04b6bca404205d99eae4c7005c700587812ed25bc99e715eb1a',APPNAME:'7766c92df82df2484eb82cd6917769f5d9ede912b3f0e8f26e3f8aaabdd357f2'}
def stop(fn,code):
    try:fn()
    except Exception as exc:
        assert code in str(exc),(code,type(exc).__name__,str(exc))
        return {'type':type(exc).__name__,'code':getattr(exc,'code',code),'message':str(exc),'cause':repr(exc.__cause__),'cleanup':str(getattr(exc,'cleanup_error',None))}
    raise AssertionError('EXPECTED STOP '+code)
def setup(label):
    F=E/label/'fixtures';O=E/label/'output';F.mkdir(parents=True);O.mkdir()
    for f in (PREVIOUS/'sandbox/fixtures').iterdir():
        if f.is_file():shutil.copyfile(f,F/f.name)
    shutil.copytree(PAIR,O/'OLD_GOOD')
    (O/'user_document.xlsx').write_bytes(b'UNRELATED USER DOCUMENT FIXTURE - PRESERVE')
    return F,O
def isolated(c,F,O):
    a=F/'ACT_template.xlsx';b=F/'APPLICATION_template.xlsx'
    return patch.multiple(c,REGISTRY=F/'DEV_registry.xlsx',ACT=a,APPLICATION=b,FIXTURES=F,OUTPUT=O,PINS={a:sha(a),b:sha(b)})
