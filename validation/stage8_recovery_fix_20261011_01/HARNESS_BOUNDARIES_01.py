"""Additional restore ambiguity, failed native verification and locking checks."""
import gc,json,os,shutil,subprocess,sys,time,traceback
from pathlib import Path
from unittest.mock import patch
from stage8_closure import core,recovery,native,storage as S,office
from support import *

OUT=RUN/sys.argv[1];OUT.mkdir(exist_ok=False);(OUT/'cases').mkdir();RESULTS=[]
PRODUCT={str(p):S.sha(p) for p in [*sorted((S.ROOT/'stage8_closure').glob('*.*')),S.ROOT/'stage8.py',S.ROOT/'STAGE8_DEV.md',S.ROOT/'Запустить Этап 8 DEV.cmd'] if p.is_file()}
S.write(OUT/'PRODUCT_BEFORE.json',PRODUCT)
def case(name,fn):
    start=time.monotonic()
    try:r={'case':name,'status':'PASS','detail':fn()}
    except Exception as e:r={'case':name,'status':'FAIL','error':str(e),'traceback':traceback.format_exc()}
    r['seconds']=round(time.monotonic()-start,3);assert all(S.sha(p)==h for p,h in PRODUCT.items())
    S.write(OUT/'cases'/(name+'.json'),r);RESULTS.append(r);save(OUT/'RESULTS.json',RESULTS)
    print(r['status'],name,r.get('error',''),flush=True)
def snap(r):return {'sha':S.sha(reg(r)) if reg(r).exists() else None,'journal':{p.name:S.sha(p) for p in (r/'journal').glob('*.json')}}
def verify_no_write(r,fn,code):
    before=snap(r);e=stopped(fn,code);assert snap(r)==before
    return {'sandbox':str(r),'stop':e,'unchanged':before}

def unknown_valid():
    r,_=partial();op=S.history(r)['active']
    # Correct native book with different balances, same target identity: requires review.
    src=S.ROOT/'validation/stage8_independent_targeted_20261011_01/EXTERNAL_VALID_4.xlsm'
    with reg(r).open('r+b') as f:f.write(src.read_bytes());f.truncate();f.flush();os.fsync(f.fileno())
    assert S.file_identity(reg(r))==op['plan']['registry_identity']
    return verify_no_write(r,lambda:recovery.preview(r),'VALID_UNRECOGNIZED_BOOK')

def corrupt_restore_stage():
    r,_=partial();p=recovery.preview(r);save(r/'RESTORE.json',p)
    w=proc('crash_restore',r/'RESTORE.json','after_restore_intent');assert w['exit']==73
    op=S.history(r)['active'];Path(op['restoration']['stage']).write_bytes(b'DAMAGED_RESTORE_STAGE')
    return verify_no_write(r,lambda:recovery.preview(r),'NAMESPACE_OR_CONTENT_AMBIGUOUS')

def native_result_failure():
    r,_=partial();p=recovery.preview(r);actual=native.inspect
    def inspect(path,*a,**k):
        if Path(path).name.startswith('r_'):raise S.Stop('INJECTED_RESTORED_NATIVE_OPEN_FAILURE')
        return actual(path,*a,**k)
    with patch('stage8_closure.recovery.native.inspect',side_effect=inspect):
        e=stopped(lambda:recovery.restore(p,recovery.confirmation(p)),'INJECTED_RESTORED_NATIVE_OPEN_FAILURE')
    assert states(r)[-1]=='RESTORE_APPLYING' and 'COMMITTED' not in states(r)
    assert S.sha(reg(r))==S.history(r)['active']['before_sha256']
    fresh=recovery.preview(r);out=recovery.restore(fresh,recovery.confirmation(fresh));assert out['status']=='RECOVERED_ABORTED'
    return {'sandbox':str(r),'safe_stop':e,'fresh_verified_resume':out,'journal':states(r)}

def parallel():
    r,_=partial();p=recovery.preview(r);plan=r/'PARALLEL_RESTORE.json';save(plan,p);marker=r/'GO'
    args=[sys.executable,'-B',str(HERE/'worker.py'),'parallel_restore',str(plan),str(marker)]
    workers=[subprocess.Popen(args,env=ENV,stdout=subprocess.PIPE,stderr=subprocess.PIPE,text=True,encoding='utf8') for _ in range(2)]
    time.sleep(.15);marker.write_text('go',encoding='utf8');results=[]
    for w in workers:
        stdout,stderr=w.communicate(timeout=100);results.append({'exit':w.returncode,'stdout':stdout,'stderr':stderr})
    assert sorted(v['exit'] for v in results)==[0,2],results
    assert states(r).count('RECOVERED_ABORTED')==1 and states(r).count('RESTORE_PREPARED')==1
    assert states(r).count('COMMITTED')==0 and balances(r)['s8-a']==12
    return {'sandbox':str(r),'workers':results,'journal':states(r)}

def writer(mode):
    r,_=partial();p=recovery.preview(r);op=S.history(r)['active'];attempts=[]
    def fault(point,target):
        if point=='after_restore_quarantine':
            path=Path(op['backup']['path']) if mode=='backup' else reg(r).parent
            name='write' if mode=='backup' else 'rename'
            v=subprocess.run([sys.executable,'-B',str(S.ROOT/'tests/stage8/work08_fix/publisher.py'),name,op['backup']['path'],str(path)],env=ENV,capture_output=True,text=True,encoding='utf8',timeout=30)
            attempts.append({'exit':v.returncode,'stdout':v.stdout,'stderr':v.stderr});assert v.returncode==2
    out=recovery.restore(p,recovery.confirmation(p),_fault=fault);assert attempts and out['status']=='RECOVERED_ABORTED'
    return {'sandbox':str(r),'attempts':attempts,'result':out}

with office.diagnostic_scope(RUN/'office'):
    case('valid_unknown_inplace_book_STOP',unknown_valid)
    case('damaged_private_restore_stage_STOP',corrupt_restore_stage)
    case('native_result_failure_STOP_then_explicit_resume',native_result_failure)
    case('two_simultaneous_restores_once',parallel)
    case('backup_writer_during_restore_blocked',lambda:writer('backup'))
    case('parent_rename_while_target_absent_blocked',lambda:writer('parent'))
S.write(OUT/'PRODUCT_AFTER.json',{p:S.sha(p) for p in PRODUCT})
print('DONE',sum(r['status']=='PASS' for r in RESULTS),sum(r['status']=='FAIL' for r in RESULTS),flush=True)
