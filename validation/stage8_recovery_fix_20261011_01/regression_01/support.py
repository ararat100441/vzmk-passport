"""Fresh artificial fixtures only; never write earlier evidence."""
import os,shutil,subprocess,sys,json
from pathlib import Path
from uuid import uuid4
from stage8_closure import core,storage as S,model

RUN=S.ROOT/'validation/stage8_recovery_fix_20261011_01'
HERE=Path(__file__).parent
ENV=os.environ|{'PYTHONPATH':str(S.ROOT),'PYTHONUTF8':'1','PYTHONDONTWRITEBYTECODE':'1'}
def save(p,v):Path(p).write_text(json.dumps(v,ensure_ascii=False,indent=2),encoding='utf8')
def reg(r):return r/'registry/УЗК_TEST.xlsm'
def req(r):return r/'inputs/CLOSE_REQUEST.json'
def fixture():
    seed=Path(S.read(RUN/'SEED.json')['sandbox'])
    r=RUN/'sandboxes'/uuid4().hex
    for name in ('registry','journal','backups','transactions','inputs'):(r/name).mkdir(parents=True)
    shutil.copyfile(reg(seed),reg(r));m=S.read(seed/'SANDBOX.json');m.pop('seal');m['sandbox_id']=r.name
    save(r/'SANDBOX.json',m|{'seal':S.digest(m)})
    for p in (seed/'inputs').glob('*.json'):
        q=S.read(p)
        if p.name=='CLOSE_REQUEST.json':
            for e in [q['shipment'],q['requirements'],q['control'],*q['documents']]:e['path']=str(r/'inputs'/Path(e['path']).name)
        save(r/'inputs'/p.name,q)
    q=S.read(req(r))
    for e in [q['shipment'],q['requirements'],q['control'],*q['documents']]:e['sha256']=S.sha(e['path'])
    save(req(r),q)
    return r
def preview(r):return core.preview(r,req(r))
def states(r):return [e['state'] for e in S.history(r)['events']]
def balances(r):return {k:v['before'] for k,v in model.registry_rows(reg(r),'STAGE8 ИСКУССТВЕННЫЙ ПРОЕКТ',{'s8-a','s8-b'}).items()}
def proc(*args):
    p=subprocess.run([sys.executable,'-B',str(HERE/'worker.py'),*map(str,args)],env=ENV,capture_output=True,text=True,encoding='utf8',timeout=100)
    return {'exit':p.returncode,'stdout':p.stdout,'stderr':p.stderr}
def partial():
    r=fixture();p=preview(r);save(r/'PLAN.json',p)
    result=proc('crash_close',r/'PLAN.json','during_publish');assert result['exit']==73,result
    op=S.history(r)['active'];S.verify_backup(op['backup'],r)
    assert S.sha(reg(r)) not in (op['before_sha256'],op['after_sha256']) and op['state']=='APPLYING'
    return r,result
def stopped(fn,contains=None):
    try:fn()
    except S.Stop as e:
        if contains:assert contains in str(e),(contains,str(e))
        return {'type':type(e).__name__,'message':str(e)}
    raise AssertionError('EXPECTED_STOP')
