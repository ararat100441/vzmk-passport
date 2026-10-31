"""Targeted Stage 8 validation only. All mutation fixtures are isolated."""
from contextlib import contextmanager
import copy
import json
import os
from pathlib import Path
import shutil
import subprocess
import sys
import time
import traceback
from unittest.mock import patch

from stage8_closure import core, native, storage as S
from stage8_closure.fixtures import create
from stage8_closure.model import registry_rows

RUN=Path(sys.argv[1]).resolve()
RESULTS=RUN/'targeted';RESULTS.mkdir(exist_ok=False)
BEFORE=core.code_bindings()
S.write(RESULTS/'PRODUCT_BEFORE.json',BEFORE)
cases=[]
env=os.environ | {'PYTHONUTF8':'1','PYTHONPATH':str(S.ROOT)}


def fixture(uzk=True):
    r=create(RUN/'sandboxes',uzk=uzk)
    return Path(r['sandbox']),Path(r['request'])


def overwrite(path,data):
    path.write_text(json.dumps(data,ensure_ascii=False,indent=2),encoding='utf8')


def amend(root,request,name,fn):
    p=root/'inputs'/name;data=S.read(p);fn(data);overwrite(p,data)
    q=S.read(request)
    for k in ('shipment','requirements','control'):
        if Path(q[k]['path'])==p:q[k]['sha256']=S.sha(p)
    for entry in q['documents']:
        if Path(entry['path'])==p:entry['sha256']=S.sha(p)
    overwrite(request,q)


def stop(function,expected):
    try:function()
    except Exception as exc:
        assert expected in str(exc),(expected,repr(exc))
        return str(exc)
    raise AssertionError('Expected STOP: '+expected)


def close(root,request,fault=None):
    p=core.preview(root,request)
    result=core.close(p,core.confirmation(p),_fault=fault)
    return p,result


def crash(root,request,point):
    plan=core.preview(root,request)
    f=root/('PREVIEW_'+point+'.json');S.write(f,plan)
    proc=subprocess.run([sys.executable,'-B',str(S.ROOT/'tests/stage8/worker.py'),'crash',str(f),point],
                        env=env,capture_output=True,text=True,encoding='utf8',timeout=35)
    assert proc.returncode==73,(proc.returncode,proc.stdout,proc.stderr)
    return plan


def fresh_recover(root):
    proc=subprocess.run([sys.executable,'-B',str(S.ROOT/'tests/stage8/worker.py'),'recover',str(root)],
                        env=env,capture_output=True,text=True,encoding='utf8',timeout=35)
    assert proc.returncode==0,(proc.stdout,proc.stderr)
    return json.loads(proc.stdout)


def case(name,fn):
    start=time.monotonic()
    print(name,'START',flush=True)
    try:
        detail=fn() or {}
        record={'name':name,'status':'PASS','detail':detail}
    except Exception as exc:
        record={'name':name,'status':'FAIL','error':repr(exc),'traceback':traceback.format_exc()}
    record['seconds']=round(time.monotonic()-start,3)
    S.write(RESULTS/(f'{len(cases)+1:02d}_'+name+'.json'),record)
    cases.append(record)
    print(name,record['status'],record.get('error',''),flush=True)


def positive():
    root,request=fixture();registry=root/'registry/УЗК_TEST.xlsm'
    initial=S.sha(registry);vba=native.vba_streams(registry)
    p,r=close(root,request)
    assert r['status']=='CLOSED_TEST' and r['official_issue'] is False
    balances=registry_rows(registry,p['payload']['identity']['project'],{'s8-a','s8-b'})
    assert balances['s8-a']['before']==9 and balances['s8-b']['before']==6
    h=S.history(root);op=h['closed'][p['payload']['key']]
    assert [e['state'] for e in h['events']]==['INTENT','BACKUP_VERIFIED','STAGED','APPLYING','COMMITTED']
    S.verify_backup(op['backup'],root);assert op['backup']['sha256']==initial
    before_inspect=S.read(root/'SANDBOX.json')['native_seed'];after=native.inspect(registry)
    assert native.vba_streams(registry)==vba
    assert after['sheets']==before_inspect['sheets']
    assert {k:v['formula'] for k,v in after['formulas'].items()}=={k:v['formula'] for k,v in before_inspect['formulas'].items()}
    assert after['formulas']['S8_FORMULAS!$A$1']['value']==23
    assert after['formulas']['S8_FORMULAS!$A$2']['value']==5
    kept=S.read(op['preservation']['path']);assert kept['all_package_parts_preserved'] is True
    committed=S.sha(registry);events=len(h['events']);backups=len(list((root/'backups').iterdir()))
    for _ in range(3):assert core.close(p,core.confirmation(p))['status']=='ALREADY_CLOSED_TEST'
    assert core.preview(root,request)['payload']['status']=='ALREADY_CLOSED_TEST'
    assert S.sha(registry)==committed and len(S.history(root)['events'])==events and len(list((root/'backups').iterdir()))==backups
    assert fresh_recover(root)['status']=='NO_PENDING_OPERATION'
    # Backup is independently reopenable by native Excel and exact SHA.
    restore=root/'BACKUP_REOPEN.xlsm';shutil.copyfile(op['backup']['path'],restore)
    assert S.sha(restore)==initial;restored=native.inspect(restore)
    S.write(RESULTS/'POSITIVE_NATIVE.json',{'before':before_inspect,'after':after,'backup_reopened':restored,'preservation':kept})
    return {'sandbox':str(root),'controlled_units':7,'balances':balances,'vba_streams':len(vba),
            'repeat_calls':3,'journal':str(root/'journal'),'backup':op['backup'],'native':'POSITIVE_NATIVE.json'}


case('successful_close_backup_debit_repeat_xlsm',positive)


def insufficient():
    root,q=fixture();before=S.sha(root/'registry/УЗК_TEST.xlsm')
    amend(root,q,'SHIPMENT.json',lambda d:d['rows'][0].update(quantity=20))
    amend(root,q,'REQUIREMENTS.json',lambda d:d['uzk'][0].update(quantity=13))
    amend(root,q,'CONTROL.json',lambda d:d['uzk'][0].update(quantity=13))
    e=stop(lambda:core.preview(root,q),'INSUFFICIENT_UZK_BALANCE')
    assert S.sha(root/'registry/УЗК_TEST.xlsm')==before and not S.history(root)['events']
    return {'sandbox':str(root),'stop':e}
case('insufficient_balance',insufficient)


def missing():
    root,q=fixture();d=S.read(q);d['documents']=[r for r in d['documents'] if r['kind']!='VIK'];overwrite(q,d)
    return {'sandbox':str(root),'stop':stop(lambda:core.preview(root,q),'MANDATORY_DOCUMENTS_MISSING')}
case('mandatory_document_missing',missing)


def no_uzk():
    root,q=fixture(False);registry=root/'registry/УЗК_TEST.xlsm';before=S.sha(registry)
    p,r=close(root,q);assert p['payload']['controlled_units']==0 and not p['payload']['deductions']
    assert r['status']=='CLOSED_TEST' and S.sha(registry)==before
    assert not list((root/'backups').iterdir()) and r['operation']['registry_mutation'] is False
    return {'sandbox':str(root),'registry_sha256':before,'registry_writes':0}
case('no_uzk_required_no_registry_write',no_uzk)


def changed_input():
    root,q=fixture();p=core.preview(root,q);file=root/'inputs/CONTROL.json';file.write_bytes(file.read_bytes()+b' ')
    before=S.sha(root/'registry/УЗК_TEST.xlsm')
    e=stop(lambda:core.close(p,core.confirmation(p)),'INPUT_SHA256_CHANGED')
    assert S.sha(root/'registry/УЗК_TEST.xlsm')==before and not S.history(root)['events']
    return {'sandbox':str(root),'stop':e}
case('input_changed_after_preview',changed_input)


def failure(point,torn=False):
    root,q=fixture();registry=root/'registry/УЗК_TEST.xlsm';before=S.sha(registry);p=core.preview(root,q)
    def fault(where,target):
        if where==point:
            if torn:target.write_bytes(b'TORN XLSM SAVE')
            else:raise OSError('INJECTED_'+point)
    expected='XLSM_PACKAGE_UNREADABLE' if torn else 'INJECTED_'+point
    e=stop(lambda:core.close(p,core.confirmation(p),_fault=fault),expected)
    assert S.sha(registry)==before
    stop(lambda:core.close(p,core.confirmation(p)),'INCOMPLETE_OPERATION')
    recovery=fresh_recover(root);assert recovery['status']=='ABORTED_UNCHANGED' and S.sha(registry)==before
    assert recovery['registry_writes']==0
    return {'sandbox':str(root),'stop':e,'recovery':recovery}
case('failure_before_save',lambda:failure('before_save'))
case('failure_at_native_save_io_boundary',lambda:failure('save_io'))
case('torn_save_result_rejected',lambda:failure('during_save',True))
case('failure_after_save_before_publish',lambda:failure('after_save'))
case('failure_before_atomic_replace',lambda:failure('before_replace'))


def process_crash(point):
    root,q=fixture();registry=root/'registry/УЗК_TEST.xlsm';before=S.sha(registry)
    p=crash(root,q,point);after=S.sha(registry)
    assert S.history(root)['active'] is not None
    result=fresh_recover(root)
    assert result['registry_writes']==0 and S.sha(registry)==after
    if point=='after_replace':
        assert result['status']=='CLOSED_TEST' and result['recovered'] is True and before!=after
        assert core.close(p,core.confirmation(p))['status']=='ALREADY_CLOSED_TEST'
    else:assert result['status']=='ABORTED_UNCHANGED' and before==after
    assert fresh_recover(root)['status']=='NO_PENDING_OPERATION'
    return {'sandbox':str(root),'recovery':result,'before':before,'after':after}
case('process_death_before_save_fresh_recovery',lambda:process_crash('after_intent'))
case('process_death_after_publish_fresh_recovery',lambda:process_crash('after_replace'))


def parallel():
    root,q=fixture();p=core.preview(root,q);f=root/'PARALLEL_PREVIEW.json';S.write(f,p)
    barrier=root/'GO';command=[sys.executable,'-B',str(S.ROOT/'tests/stage8/worker.py'),'parallel',str(f),str(barrier)]
    jobs=[subprocess.Popen(command,env=env,stdout=subprocess.PIPE,stderr=subprocess.PIPE,text=True,encoding='utf8') for _ in range(2)]
    barrier.write_text('GO')
    outputs=[]
    for job in jobs:
        out,err=job.communicate(timeout=35);assert not err,err;outputs.append(json.loads(out))
    assert sorted(r['status'] for r in outputs)==['CLOSED_TEST','STOP'],outputs
    assert 'PARALLEL_OPERATION' in next(r['error'] for r in outputs if r['status']=='STOP')
    assert len(S.history(root)['closed'])==1
    balances=registry_rows(root/'registry/УЗК_TEST.xlsm',p['payload']['identity']['project'],{'s8-a','s8-b'})
    assert balances['s8-a']['before']==9 and balances['s8-b']['before']==6
    return {'sandbox':str(root),'statuses':[{'status':r['status'],'error':r.get('error')} for r in outputs],'balances':balances}
case('two_simultaneous_processes',parallel)


def denied_confirmation():
    root,q=fixture();p=core.preview(root,q)
    e=stop(lambda:core.close(p,''),'EXPLICIT_CONFIRMATION_REQUIRED')
    assert not S.history(root)['events']
    return {'sandbox':str(root),'stop':e}
case('explicit_confirmation_required',denied_confirmation)


def unconfirmed():
    root,q=fixture();amend(root,q,'CONTROL.json',lambda d:d.update(all_required_control_confirmed=False))
    return {'sandbox':str(root),'stop':stop(lambda:core.preview(root,q),'CONTROL_NOT_CONFIRMED')}
case('actual_control_not_confirmed',unconfirmed)


def sample_doc():
    root,q=fixture();amend(root,q,'UZK.json',lambda d:d.update(status='SAMPLE_ONLY',results_confirmed=False))
    return {'sandbox':str(root),'stop':stop(lambda:core.preview(root,q),'DOCUMENT_NOT_CONFIRMED')}
case('sample_only_document_rejected',sample_doc)


def dev_passport():
    root,q=fixture();d=S.read(q);d['passport']='9/2026';overwrite(q,d)
    before=S.sha(root/'registry/УЗК_TEST.xlsm')
    e=stop(lambda:core.preview(root,q),'PASSPORT_9_2026_DEV_TEST_FORBIDDEN')
    d['mode']='OFFICIAL';d['official_issue']=True;overwrite(q,d)
    stop(lambda:core.preview(root,q),'OFFICIAL_CLOSURE_FORBIDDEN')
    assert S.sha(root/'registry/УЗК_TEST.xlsm')==before and not S.history(root)['events']
    return {'sandbox':str(root),'stop':e,'official_promotion':'REJECTED'}
case('passport_9_2026_and_official_promotion_forbidden',dev_passport)


def registry_drift():
    root,q=fixture();p=core.preview(root,q);registry=root/'registry/УЗК_TEST.xlsm';registry.write_bytes(registry.read_bytes()+b'EXTERNAL_CHANGE')
    current=S.sha(registry);e=stop(lambda:core.close(p,core.confirmation(p)),'REGISTRY_CHANGED_OR_JOURNAL_LOST')
    assert S.sha(registry)==current and not S.history(root)['events']
    return {'sandbox':str(root),'stop':e}
case('registry_changed_after_preview',registry_drift)


def damaged_backup():
    root,q=fixture();p=core.preview(root,q);registry=root/'registry/УЗК_TEST.xlsm';before=S.sha(registry)
    def fault(point,path):
        if point=='after_backup':path.write_bytes(b'DAMAGED BACKUP')
    e=stop(lambda:core.close(p,core.confirmation(p),_fault=fault),'BACKUP_DAMAGED')
    assert S.sha(registry)==before
    return {'sandbox':str(root),'stop':e,'recovery':fresh_recover(root)}
case('backup_integrity_failure_forbids_save',damaged_backup)


def backup_io_failure():
    root,q=fixture();p=core.preview(root,q);registry=root/'registry/УЗК_TEST.xlsm';before=S.sha(registry)
    with patch.object(S,'backup',side_effect=OSError('SIMULATED_BACKUP_DISK_FULL')):
        e=stop(lambda:core.close(p,core.confirmation(p)),'SIMULATED_BACKUP_DISK_FULL')
    assert S.sha(registry)==before
    return {'sandbox':str(root),'stop':e,'recovery':fresh_recover(root)}
case('backup_io_failure_forbids_write',backup_io_failure)


def ambiguous():
    root,q=fixture();p=crash(root,q,'after_replace');registry=root/'registry/УЗК_TEST.xlsm'
    registry.write_bytes(registry.read_bytes()+b'UNKNOWN STATE');before=S.sha(registry)
    e=stop(lambda:core.recover(root),'AMBIGUOUS_OPERATION_STOP_NO_REPLAY')
    assert S.sha(registry)==before and S.history(root)['active'] is not None
    return {'sandbox':str(root),'stop':e,'registry_writes_on_recovery':0}
case('ambiguous_recovery_stop_no_replay',ambiguous)


def backup_corrupt_recovery():
    root,q=fixture();crash(root,q,'after_replace');op=S.history(root)['active']
    Path(op['backup']['path']).write_bytes(b'CORRUPT');before=S.sha(root/'registry/УЗК_TEST.xlsm')
    e=stop(lambda:core.recover(root),'BACKUP_DAMAGED')
    assert S.sha(root/'registry/УЗК_TEST.xlsm')==before
    return {'sandbox':str(root),'stop':e}
case('damaged_backup_blocks_recovery',backup_corrupt_recovery)


def journal_corrupt():
    root,q=fixture();close(root,q);last=sorted((root/'journal').glob('*.json'))[-1]
    d=S.read(last);d['closed_status']='UNEXPLAINED';overwrite(last,d)
    e=stop(lambda:core.preview(root,q),'JOURNAL_EVENT_DAMAGED')
    return {'sandbox':str(root),'stop':e}
case('journal_corruption_blocks_reopen',journal_corrupt)


def lost_journal():
    root,q=fixture();close(root,q);archive=root/'REMOVED_JOURNAL';archive.mkdir()
    for p in (root/'journal').glob('*.json'):p.rename(archive/p.name)
    e=stop(lambda:core.preview(root,q),'REGISTRY_CHANGED_OR_JOURNAL_LOST')
    return {'sandbox':str(root),'stop':e}
case('lost_journal_cannot_repeat_debit',lost_journal)


def external_lock():
    import win32file,win32con
    root,q=fixture();p=core.preview(root,q);registry=root/'registry/УЗК_TEST.xlsm'
    handle=win32file.CreateFile(str(registry),win32con.GENERIC_READ|win32con.GENERIC_WRITE,
        win32con.FILE_SHARE_READ|win32con.FILE_SHARE_WRITE,None,win32con.OPEN_EXISTING,0,None)
    try:e=stop(lambda:core.close(p,core.confirmation(p)),'INPUT_LOCKED')
    finally:handle.Close()
    assert not S.history(root)['events']
    return {'sandbox':str(root),'stop':e}
case('external_registry_writer_is_denied',external_lock)


def zero_remaining():
    root,q=fixture()
    amend(root,q,'SHIPMENT.json',lambda d:d['rows'][0].update(quantity=12))
    amend(root,q,'REQUIREMENTS.json',lambda d:d['uzk'][0].update(quantity=12))
    amend(root,q,'CONTROL.json',lambda d:d['uzk'][0].update(quantity=12))
    p,r=close(root,q)
    balance=registry_rows(root/'registry/УЗК_TEST.xlsm',p['payload']['identity']['project'],{'s8-a'})
    assert balance['s8-a']['before']==0 and r['status']=='CLOSED_TEST'
    return {'sandbox':str(root),'balance':balance}
case('zero_balance_is_preserved_not_blank_one',zero_remaining)


def previous_backups():
    root,q=fixture();p,r=close(root,q);first=r['operation']['backup'];first_sha=S.sha(first['path'])
    d=S.read(q);d['passport']='8002/2099';overwrite(q,d)
    for name in ['SHIPMENT.json','REQUIREMENTS.json','CONTROL.json','QUALITY.json','APPENDIX_1.json','VIK.json','AKZ.json','AKZ_APPENDIX.json','UZK.json']:
        amend(root,q,name,lambda x:x.update(passport='8002/2099'))
    p2,r2=close(root,q)
    assert r2['status']=='CLOSED_TEST' and len(S.history(root)['closed'])==2
    assert len(list((root/'backups').iterdir()))==2 and S.sha(first['path'])==first_sha
    balances=registry_rows(root/'registry/УЗК_TEST.xlsm',p['payload']['identity']['project'],{'s8-a','s8-b'})
    assert balances['s8-a']['before']==6 and balances['s8-b']['before']==2
    assert core.close(p,core.confirmation(p))['status']=='ALREADY_CLOSED_TEST'
    return {'sandbox':str(root),'backups_kept':2,'balances':balances}
case('previous_backups_retained_across_passports',previous_backups)


def boundary():
    e=stop(lambda:S.sandbox(S.ROOT/'DEV'),'ISOLATED_STAGE8_DIRECTORY_REQUIRED')
    root,q=fixture();d=S.read(q);d['control']['path']=str(S.ROOT/'DEV/DEV_UZK_CHECKPOINT.json');overwrite(q,d)
    e2=stop(lambda:core.preview(root,q),'INPUT_OUTSIDE_SANDBOX')
    return {'production_and_dev_write_boundary':e,'external_reference':e2}
case('write_boundary_and_input_path_escape',boundary)


def stage7():
    files=sorted((S.ROOT/'validation').glob('stage7_*/**/STAGE7_SELECTED_TEST_MANIFEST.json'))
    assert files,'Existing Stage 7 manifest required'
    results=[core.inspect_stage7(p) | {'manifest':str(p),'sha256':S.sha(p)} for p in files]
    assert all(r['status']=='BLOCKED_DEV_TEST' and r['registry_writes']==0 for r in results)
    S.write(RESULTS/'STAGE7_BLOCKED.json',results)
    return {'manifests_read':len(results),'status':'BLOCKED_DEV_TEST','registry_writes':0}
case('existing_stage7_9_2026_read_only_block',stage7)


AFTER=core.code_bindings();assert BEFORE==AFTER,'Tested code changed during validation'
S.write(RESULTS/'PRODUCT_AFTER.json',AFTER)
S.write(RESULTS/'RESULTS.json',{'cases':cases,'passed':sum(c['status']=='PASS' for c in cases),
    'failed':sum(c['status']=='FAIL' for c in cases),'scope':'STAGE8_ONLY_ISOLATED_TEST',
    'product_sha256_equal':BEFORE==AFTER})
print(json.dumps({'passed':sum(c['status']=='PASS' for c in cases),'failed':sum(c['status']=='FAIL' for c in cases)},ensure_ascii=False),flush=True)
sys.exit(1 if any(c['status']=='FAIL' for c in cases) else 0)
