"""Bounded recovery acceptance with actual Excel, NTFS and separate processes."""
import copy,gc,json,os,runpy,shutil,subprocess,sys,time,traceback
from pathlib import Path
from unittest.mock import patch
import tkinter as tk
import win32file,win32con
from stage8_closure import core,storage as S,native,recovery,office
from stage8_closure.gui import App
from support import *

BATCH=RUN/sys.argv[1];BATCH.mkdir(exist_ok=False);(BATCH/'cases').mkdir()
ONLY=set(sys.argv[2:]);RESULTS=[]
PRODUCT={str(p):S.sha(p) for p in [*sorted((S.ROOT/'stage8_closure').glob('*.*')),S.ROOT/'stage8.py',S.ROOT/'STAGE8_DEV.md',S.ROOT/'Запустить Этап 8 DEV.cmd'] if p.is_file()}
S.write(BATCH/'PRODUCT_BEFORE.json',PRODUCT)
def case(name,fn):
    if ONLY and name not in ONLY:return
    start=time.monotonic()
    try:r={'case':name,'status':'PASS','detail':fn(),'seconds':round(time.monotonic()-start,3)}
    except Exception as e:r={'case':name,'status':'FAIL','error':str(e),'traceback':traceback.format_exc()}
    assert all(S.sha(p)==h for p,h in PRODUCT.items()),'PRODUCT_CHANGED'
    S.write(BATCH/'cases'/(name+'.json'),r);RESULTS.append(r);save(BATCH/'RESULTS.json',RESULTS)
    print(r['status'],name,r.get('error',''),flush=True)

def restored(r):
    p=recovery.preview(r);return recovery.restore(p,recovery.confirmation(p))
def unchanged_snapshot(r):
    return {'target':S.sha(reg(r)) if reg(r).exists() else None,
            'journal':{p.name:S.sha(p) for p in (r/'journal').glob('*.json')}}
def refuse_without_change(r,fn,code=None):
    old=unchanged_snapshot(r);e=stopped(fn,code);assert unchanged_snapshot(r)==old
    return {'sandbox':str(r),'stop':e,'unchanged':old}

def success():
    r,c=partial();p=recovery.preview(r);original=S.history(r)['active'];old=reg(r).read_bytes()
    assert 'COMMITTED' not in states(r)
    pre=stopped(lambda:core.recover(r),'AMBIGUOUS_OPERATION')
    result=recovery.restore(p,recovery.confirmation(p))
    assert result['status']=='RECOVERED_ABORTED' and result['decrements']==0 and result['closed'] is False
    assert result['old_identity']!=result['new_identity'] and S.file_identity(reg(r))==result['new_identity']
    assert S.sha(reg(r))==original['before_sha256']==S.sha(original['backup']['path'])
    h=S.history(r);assert h['active'] is None and not h['closed'] and len(h['restored'])==1
    assert Path(h['restored'][0]['restoration']['quarantine']).read_bytes()==old
    fresh=proc('recover',r);assert fresh['exit']==0 and json.loads(fresh['stdout'])['status']=='NO_PENDING_OPERATION'
    assert states(r).count('COMMITTED')==0 and balances(r)=={'s8-a':12,'s8-b':10}
    repeat=refuse_without_change(r,lambda:recovery.restore(p,recovery.confirmation(p)),'NO_PENDING')
    stale=stopped(lambda:core.close(S.read(r/'PLAN.json'),core.confirmation(S.read(r/'PLAN.json'))),'FILE_ID_CHANGED')
    n=native.inspect(reg(r),{'УЗК!C24':12,'УЗК!C25':10})
    assert native.vba_streams(reg(r))==native.vba_streams(original['backup']['path'])
    assert native.package(reg(r))==native.package(original['backup']['path'])
    new=preview(r);out=core.close(new,core.confirmation(new));assert out['status']=='CLOSED_TEST'
    final=S.sha(reg(r));again=core.close(new,core.confirmation(new));assert again['status']=='ALREADY_CLOSED_TEST'
    assert S.sha(reg(r))==final and states(r).count('COMMITTED')==1 and balances(r)['s8-a']==9
    return {'sandbox':str(r),'crash':c,'prior_safe_stop':pre,'restore':result,'fresh_reopen':fresh,
            'repeat_restore':repeat,'old_preview_stop':stale,'restored_native':n,'full_package_equals_backup':True,
            'new_explicit_operation':out['operation']['operation_id'],'initial_failed_operation':original['operation_id'],
            'journal':states(r),'final':balances(r)}

def identity_after_restore():
    r,_=partial();result=restored(r);data=reg(r).read_bytes();incoming=r/'transactions/unknown.xlsm'
    incoming.write_bytes(data);os.replace(incoming,reg(r))
    e=refuse_without_change(r,lambda:preview(r),'FILE_ID_CHANGED')
    e['official_new_identity']=result['new_identity'];e['unknown_identity']=S.file_identity(reg(r))
    return e

def crash_restore(point):
    r,_=partial();p=recovery.preview(r);save(r/'RESTORE.json',p);before=S.history(r)['active']['backup']
    worker=proc('crash_restore',r/'RESTORE.json',point);assert worker['exit']==73,worker
    op=S.history(r)['active'];S.verify_backup(before,r);assert 'COMMITTED' not in states(r)
    blocked=proc('preview',r);assert blocked['exit']==2,blocked
    p2=recovery.preview(r);result=recovery.restore(p2,recovery.confirmation(p2))
    assert result['status']=='RECOVERED_ABORTED' and balances(r)['s8-a']==12 and 'COMMITTED' not in states(r)
    return {'sandbox':str(r),'point':point,'worker':worker,'pending_state':op['state'],
            'resume_phase':p2['payload']['phase'],'new_preview_required':p2['seal']!=p['seal'] if point!='restore_copy_chunk' else True,
            'blocked_close':blocked,'backup':before,'result':result,'journal':states(r)}

def bad_backup(mode):
    r,_=partial();p=recovery.preview(r);op=S.history(r)['active'];b=Path(op['backup']['path'])
    if mode=='damaged':b.write_bytes(b'BROKEN_BACKUP')
    else:shutil.copyfile(r/'transactions'/op['operation_id']/'stage.xlsm',b)
    a=refuse_without_change(r,lambda:recovery.restore(p,recovery.confirmation(p)),'BACKUP_DAMAGED')
    a['preview_stop']=stopped(lambda:recovery.preview(r),'BACKUP_DAMAGED');return a

def damaged_journal():
    r,_=partial();p=recovery.preview(r);(r/'journal/00000001.json').write_bytes(b'{BROKEN')
    return refuse_without_change(r,lambda:recovery.restore(p,recovery.confirmation(p)),'UNREADABLE_JSON')

def manual_backup_replace():
    r,_=partial();op=S.history(r)['active'];incoming=r/'transactions/manual.xlsm';shutil.copyfile(op['backup']['path'],incoming)
    os.replace(incoming,reg(r))
    e=refuse_without_change(r,lambda:recovery.preview(r),'UNKNOWN_EXTERNAL_REGISTRY_IDENTITY')
    e['normal_recover_stop']=stopped(lambda:core.recover(r),'FILE_ID');return e

def no_confirmation():
    r,_=partial();p=recovery.preview(r)
    return refuse_without_change(r,lambda:recovery.restore(p,''),'EXPLICIT_CONFIRMATION')

def stale_preview():
    r,_=partial();p=recovery.preview(r)
    with reg(r).open('ab') as f:f.write(b'EXTERNAL_CHANGE')
    return refuse_without_change(r,lambda:recovery.restore(p,recovery.confirmation(p)),'PREVIEW_STALE')

def full_after():
    r=fixture();p=preview(r);save(r/'PLAN.json',p);w=proc('crash_close',r/'PLAN.json','after_replace');assert w['exit']==73
    e=refuse_without_change(r,lambda:recovery.preview(r),'FULL_RESULT_USE_JOURNAL_RECOVER')
    before=S.sha(reg(r));result=core.recover(r);assert result['status']=='CLOSED_TEST' and result['registry_writes']==0
    assert S.sha(reg(r))==before and balances(r)['s8-a']==9 and states(r).count('COMMITTED')==1
    return e|{'journal_recover':result,'journal':states(r)}

def committed():
    r=fixture();p=preview(r);out=core.close(p,core.confirmation(p))
    return refuse_without_change(r,lambda:recovery.preview(r),'NO_PENDING')|{'closed_operation':out['operation']['operation_id']}

def change_passport(r,passport):
    for f in (r/'inputs').glob('*.json'):
        q=S.read(f);q['passport']=passport;save(f,q)
    q=S.read(req(r))
    for e in [q['shipment'],q['requirements'],q['control'],*q['documents']]:e['sha256']=S.sha(e['path'])
    save(req(r),q)

def later_closures():
    r,_=partial();rp=recovery.preview(r);first=restored(r);p=preview(r);core.close(p,core.confirmation(p))
    change_passport(r,'8002/2099');p2=preview(r);core.close(p2,core.confirmation(p2))
    assert balances(r)['s8-a']==6 and states(r).count('COMMITTED')==2
    e=refuse_without_change(r,lambda:recovery.restore(rp,recovery.confirmation(rp)),'NO_PENDING')
    change_passport(r,'8003/2099')
    # Only 2 S8-B remain after two closes; this third fixture deliberately
    # requests one, so the acceptance reaches publication rather than balance STOP.
    for name in ('REQUIREMENTS.json','CONTROL.json'):
        f=r/'inputs'/name;q=S.read(f);q['uzk'][1]['quantity']=1;save(f,q)
    q=S.read(req(r))
    for e in [q['shipment'],q['requirements'],q['control'],*q['documents']]:e['sha256']=S.sha(e['path'])
    save(req(r),q);p3=preview(r);save(r/'PLAN3.json',p3)
    w=proc('crash_close',r/'PLAN3.json','during_publish');assert w['exit']==73
    last=restored(r);assert balances(r)['s8-a']==6 and len(S.history(r)['closed'])==2
    assert first['new_identity']!=last['new_identity'] and last['decrements']==0
    return {'sandbox':str(r),'old_restore_refused':e,'first_restore':first,'later_restore':last,'final_balances':balances(r),'journal':states(r)}

def publisher(mode,src,dst):
    p=subprocess.run([sys.executable,'-B',str(S.ROOT/'tests/stage8/work08_fix/publisher.py'),mode,str(src),str(dst)],env=ENV,capture_output=True,text=True,encoding='utf8',timeout=30)
    return {'exit':p.returncode,'stdout':p.stdout,'stderr':p.stderr}

def race(mode,point):
    r,_=partial();p=recovery.preview(r);op=S.history(r)['active'];incoming=r/'transactions/external.xlsm'
    shutil.copyfile(op['backup']['path'],incoming);attempts=[]
    def fault(phase,target):
        if phase==point:
            a=publisher(mode,incoming,reg(r));attempts.append(a);assert a['exit']==2,a
    result=recovery.restore(p,recovery.confirmation(p),_fault=fault)
    assert attempts and result['status']=='RECOVERED_ABORTED' and balances(r)['s8-a']==12
    return {'sandbox':str(r),'external_process':attempts,'result':result,'journal':states(r)}

def gap_creator():
    r,_=partial();p=recovery.preview(r);op=S.history(r)['active'];incoming=r/'transactions/external.xlsm'
    shutil.copyfile(op['backup']['path'],incoming);attempt=[]
    def fault(point,target):
        if point=='after_restore_quarantine':
            v=publisher('replace',incoming,reg(r));attempt.append(v);assert v['exit']==0
    error=stopped(lambda:recovery.restore(p,recovery.confirmation(p),_fault=fault),'NO_REPLACE_FAILED')
    assert S.sha(reg(r))==op['before_sha256'] and 'COMMITTED' not in states(r)
    current=unchanged_snapshot(r);again=stopped(lambda:recovery.preview(r),'AMBIGUOUS');assert current==unchanged_snapshot(r)
    assert S.history(r)['active']['state']=='RESTORE_APPLYING'
    return {'sandbox':str(r),'creator':attempt,'stop':error,'fresh_stop':again,'external_file_preserved':True,'journal':states(r)}

def existing_writer():
    r,_=partial();p=recovery.preview(r)
    h=win32file.CreateFile(str(reg(r)),win32con.GENERIC_WRITE,7,None,win32con.OPEN_EXISTING,0,None)
    try:return refuse_without_change(r,lambda:recovery.restore(p,recovery.confirmation(p)),'GUARD_OR_IO_FAILURE')
    finally:h.Close()

def corrupt_restored_proof():
    r,_=partial();out=restored(r);Path(out['native_evidence']).write_bytes(b'DAMAGED')
    return refuse_without_change(r,lambda:preview(r),'RESTORED_NATIVE_EVIDENCE_DAMAGED')

def gui_flow():
    r,_=partial();window=tk.Tk();window.withdraw();app=App(window);questions=[];errors=[]
    try:
        app.load(req(r));window.update_idletasks();text=app.text.get('1.0','end')
        for label in ('Обнаружена незавершённая операция','Статус:','Причина:','Проверенный backup:','SHA-256 backup:','Затрагиваемый реестр:','списаний 0'):
            assert label in text,label
        assert str(app.close_button['state'])=='disabled' and str(app.restore_button['state'])=='normal'
        before=unchanged_snapshot(r)
        with patch('stage8_closure.gui.messagebox.askyesno',side_effect=lambda *a,**k:questions.append({'args':a,'kwargs':k}) or False):app.restore_button.invoke()
        assert questions[0]['kwargs']['default']=='no' and unchanged_snapshot(r)==before
        with patch('stage8_closure.gui.messagebox.askyesno',return_value=True),patch('stage8_closure.gui.messagebox.showerror',side_effect=lambda *a:errors.append(a)):
            app.restore_button.invoke()
        assert not errors,errors
        success_text=app.text.get('1.0','end');assert 'RECOVERED_ABORTED' in success_text and 'Списаний: 0' in success_text
        assert str(app.restore_button['state'])=='disabled' and str(app.close_button['state'])=='disabled'
        # New program/widget state: a new preview is required; no retained confirmation.
        window.destroy();window=tk.Tk();window.withdraw();app=App(window);app.load(req(r))
        assert str(app.close_button['state'])=='normal' and not app.confirm.get()
        return {'sandbox':str(r),'pending_text':text,'separate_confirmation':questions,'success_text':success_text,
                'new_app_new_preview':True,'real_tk_widgets':True,'human_visual_E2E':'NOT_RUN'}
    finally:window.destroy()

def gui_stop():
    r,_=partial();op=S.history(r)['active'];Path(op['backup']['path']).write_bytes(b'BROKEN')
    w=tk.Tk();w.withdraw();app=App(w)
    try:
        app.load(req(r));text=app.text.get('1.0','end')
        assert 'BACKUP_DAMAGED' in text and str(app.restore_button['state'])=='disabled'
        assert str(app.close_button['state'])=='disabled'
        return {'sandbox':str(r),'text':text,'real_tk_widgets':True}
    finally:w.destroy()

def cli():
    r,_=partial();path=r/'CLI_RESTORE.json'
    def command(*args):
        v=subprocess.run([sys.executable,'-B',str(S.ROOT/'stage8.py'),*map(str,args)],env=ENV,capture_output=True,text=True,encoding='utf8',timeout=100)
        return {'exit':v.returncode,'stdout':v.stdout,'stderr':v.stderr}
    before=unchanged_snapshot(r);p=command('restore-preview',r,'--save-preview',path);assert p['exit']==0,p
    assert before==unchanged_snapshot(r)
    plan=S.read(path);bad=command('restore',path,'--confirm','WRONG');assert bad['exit']==2 and before==unchanged_snapshot(r)
    good=command('restore',path,'--confirm',recovery.confirmation(plan));assert good['exit']==0,good
    assert balances(r)['s8-a']==12 and states(r).count('COMMITTED')==0
    return {'sandbox':str(r),'preview':p,'wrong_confirmation':bad,'restore':good}

with office.diagnostic_scope(RUN/'office'):
    case('partial_XLSM_supported_operator_recovery',success)
    case('unknown_identity_after_official_restore_STOP',identity_after_restore)
    for phase in ('restore_copy_chunk','after_restore_intent','before_restore_quarantine','after_restore_quarantine','after_restore_publish','before_restore_terminal'):
        case('crash_'+phase,lambda phase=phase:crash_restore(phase))
    case('damaged_backup_STOP',lambda:bad_backup('damaged'))
    case('substituted_valid_backup_STOP',lambda:bad_backup('substituted'))
    case('damaged_journal_STOP',damaged_journal)
    case('manual_backup_replacement_still_STOP',manual_backup_replace)
    case('explicit_confirmation_required',no_confirmation)
    case('stale_restore_preview_STOP',stale_preview)
    case('full_after_finalize_only_no_rollback',full_after)
    case('COMMITTED_old_backup_restore_forbidden',committed)
    case('later_COMMITTED_preserved_across_restore',later_closures)
    case('external_write_before_quarantine_blocked',lambda:race('write','before_restore_quarantine'))
    case('external_native_replace_before_quarantine_blocked',lambda:race('native','before_restore_quarantine'))
    case('external_native_replace_after_publish_blocked',lambda:race('native','after_restore_publish'))
    case('external_creator_in_name_gap_preserved_STOP',gap_creator)
    case('existing_writer_restore_STOP',existing_writer)
    case('restored_native_evidence_tamper_STOP',corrupt_restored_proof)
    case('GUI_real_widgets_cancel_restore_restart',gui_flow)
    case('GUI_backup_STOP_disables_restore',gui_stop)
    case('CLI_separate_preview_confirm_restore',cli)
S.write(BATCH/'PRODUCT_AFTER.json',{p:S.sha(p) for p in PRODUCT})
print('DONE',sum(r['status']=='PASS' for r in RESULTS),sum(r['status']=='FAIL' for r in RESULTS),flush=True)
