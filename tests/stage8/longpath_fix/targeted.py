"""New isolated fixtures; never execute a harness in its historical directory."""
import gc, hashlib, json, os, shutil, subprocess, sys, time, traceback
from contextlib import contextmanager
from pathlib import Path
from unittest.mock import patch
from stage8_closure import core, native, recovery, shortcopy, storage as S, office

RUN=S.ROOT/'validation/stage8_longpath_fix_20261011_001'
sys.path.insert(0,str(RUN))
import review as R
H=R.H
ROOTS=S.read(RUN/'TEST_ROOTS.json')
os.environ['STAGE8_REVIEW_FIXTURE_PARENT']=str(Path(ROOTS['long'])/'sandboxes')
H.SEED=Path(S.read(RUN/'SEED.json')['sandbox']);H.EXTERNAL=RUN/'EXTERNAL_VALID_4.xlsm'
OUT=RUN/sys.argv[1];OUT.mkdir(exist_ok=False);(OUT/'cases').mkdir()
PRODUCT=S.read(RUN/'TEST_PRODUCT.json');RESULTS=[]

def case(name,fn):
    print('START',name,flush=True);t=time.monotonic()
    try:
        detail=fn();gc.collect();item={'name':name,'status':'PASS','detail':detail}
    except Exception as e:item={'name':name,'status':'FAIL','error':str(e),'traceback':traceback.format_exc()}
    item['seconds']=round(time.monotonic()-t,3);RESULTS.append(item)
    S.write(OUT/'cases'/(name+'.json'),item)
    (OUT/'RESULTS.json').write_text(json.dumps(RESULTS,ensure_ascii=False,indent=2),encoding='utf8')
    assert {p:S.sha(p) for p in PRODUCT}==PRODUCT,'PRODUCT_CHANGED'
    print(item['status'],name,item.get('error',''),flush=True)

def invalid(mode,restore=False,resume=False):
    if restore:
        r,_=R.partial();p=R.prest(r)
        if resume:
            R.save(r/'RESTORE.json',p);assert R.proc('restore_crash',r/'RESTORE.json','after_restore_intent')['exit']==73
            p=R.prest(r)
        fn=lambda:recovery.restore(p,recovery.confirmation(p))
    else:
        r=R.fixture();p=R.pclose(r);fn=lambda:core.close(p,core.confirmation(p))
    before=R.snap(r)
    base=Path(ROOTS['short'])/('invalid_'+mode+'_'+OUT.name)
    if mode=='file':base.write_bytes(b'NOT_A_DIRECTORY')
    elif mode=='long':base=RUN/('q'*90);base.mkdir(exist_ok=True)
    elif mode=='missing':pass
    elif mode=='acl':
        import win32api,win32con,win32security
        base.mkdir(exist_ok=False)
        old=win32security.GetNamedSecurityInfo(str(base),win32security.SE_FILE_OBJECT,win32security.DACL_SECURITY_INFORMATION)
        token=win32security.OpenProcessToken(win32api.GetCurrentProcess(),win32con.TOKEN_QUERY)
        sid=win32security.GetTokenInformation(token,win32security.TokenUser)[0];token.Close()
        dacl=win32security.ACL();dacl.AddAccessDeniedAce(win32security.ACL_REVISION,4,sid)
        dacl.AddAccessAllowedAce(win32security.ACL_REVISION,0x1f01ff,sid)
        win32security.SetNamedSecurityInfo(str(base),win32security.SE_FILE_OBJECT,win32security.DACL_SECURITY_INFORMATION|win32security.PROTECTED_DACL_SECURITY_INFORMATION,None,None,dacl,None)
    try:
        with patch.object(shortcopy,'BASE',base):
            # The shared namespace guard labels the actual Win32 CreateDirectory
            # access denial; unavailable roots have the explicit native label.
            e=H.stopped(fn,'GUARD_OR_IO_FAILURE' if mode=='acl' else 'NATIVE_SHORT_WORKSPACE')
            if restore:pe=H.stopped(lambda:R.prest(r),'NATIVE_SHORT_WORKSPACE')
        assert R.snap(r)==before
        assert 'COMMITTED' not in R.states(r)
        if not restore:assert not R.states(r) and not list((r/'backups').iterdir())
        return {'sandbox':str(r),'mode':mode,'error':e,'before_equals_after':True,'before_target_mutation':True}
    finally:
        if mode=='acl':win32security.SetNamedSecurityInfo(str(base),win32security.SE_FILE_OBJECT,win32security.DACL_SECURITY_INFORMATION|win32security.UNPROTECTED_DACL_SECURITY_INFORMATION,None,None,old.GetSecurityDescriptorDacl(),None)

def locks(which):
    r=R.fixture();source=R.reg(r);before=S.sha(source);identity=S.file_identity(source)
    actual=native._inspect;attempts=[];seen=[]
    def inspect(copy,*a,**kw):
        seen.append(copy)
        import win32security
        sd=win32security.GetNamedSecurityInfo(str(copy.parent),win32security.SE_FILE_OBJECT,win32security.DACL_SECURITY_INFORMATION)
        assert sd.GetSecurityDescriptorControl()[0]&0x1000,'DACL_NOT_PROTECTED'
        for mode in ('write','replace','native','parent'):
            incoming=r/'transactions'/('incoming_'+mode+'.xlsm');shutil.copyfile(H.EXTERNAL,incoming)
            target=copy if which=='copy' else source
            if mode=='parent':target=target.parent
            v=H.proc('publisher',mode,incoming,target)
            assert v['exit']==2,v;attempts.append(v)
        return actual(copy,*a,**kw)
    with patch.object(native,'_inspect',side_effect=inspect):info=native.inspect(source,diagnostics=OUT/'office')
    b=info['verified_copy'];assert b['source_and_copy_rechecked'] and b['source_sha256']==before
    assert b['source_path']==str(source) and b['source_identity']==identity
    assert S.sha(source)==before and S.file_identity(source)==identity and not seen[0].exists() and not seen[0].parent.exists()
    return {'sandbox':str(r),'which':which,'actual_external_process_attempts':attempts,'native':info,'cleaned':True}

def tamper(kind):
    r=R.fixture();source=R.reg(r);before=S.sha(source);real=S.frozen;seen=[]
    @contextmanager
    def freeze(paths):
        for p in map(Path,paths):
            if p.name=='v.xlsm':
                seen.append(p)
                if kind=='content':p.write_bytes(b'FOREIGN_TEST_BYTES')
                else:
                    q=p.with_name('incoming.xlsm');q.write_bytes(b'FOREIGN_TEST_BYTES');os.replace(q,p)
        with real(paths):yield
    try:
        with patch.object(S,'frozen',side_effect=freeze),patch('win32com.client.DispatchEx') as excel:
            error=H.stopped(lambda:native.inspect(source),'NATIVE_SHORT_COPY_CHANGED')
            assert not excel.called
        assert S.sha(source)==before and not R.states(r)
        if kind=='identity':assert seen[0].read_bytes()==b'FOREIGN_TEST_BYTES'
        return {'error':error,'source_unchanged':True,'Excel_not_started':True,'foreign_copy_preserved':kind=='identity'}
    finally:
        # This is the test's known injected file, never an arbitrary orphan.
        if seen and seen[0].exists():seen[0].unlink();seen[0].parent.rmdir()

def exception_cleanup():
    r=R.fixture();bad=r/'transactions/bad.xlsm';bad.write_bytes(b'PARTIAL_XLSM')
    real=shortcopy.workspace;seen=[]
    @contextmanager
    def work():
        with real() as w:seen.append(w.path);yield w
    with patch.object(native,'workspace',side_effect=work):e=H.stopped(lambda:native.inspect(bad),'XLSM_PACKAGE_UNREADABLE')
    assert seen and all(not p.exists() for p in seen),seen
    return {'error':e,'temporary_data_cleaned':True}

def gui(unavailable=False):
    import tkinter as tk
    from stage8_closure.gui import App
    r,_=R.partial();window=tk.Tk();window.withdraw();app=App(window);errors=[]
    try:
        if unavailable:
            with patch.object(shortcopy,'BASE',Path(ROOTS['short'])/'missing_gui'):app.load(R.req(r))
            text=app.text.get('1.0','end');assert 'NATIVE_SHORT_WORKSPACE' in text
            assert str(app.restore_button['state'])=='disabled' and str(app.close_button['state'])=='disabled'
            return {'sandbox':str(r),'STOP_text':text,'real_Tk_widgets':True,'visual_E2E':'NOT_RUN: native computer-control API unavailable'}
        app.load(R.req(r));assert str(app.restore_button['state'])=='normal'
        text=app.text.get('1.0','end');before=R.snap(r);questions=[]
        with patch('stage8_closure.gui.messagebox.askyesno',side_effect=lambda *a,**kw:questions.append(kw) or False):app.restore_button.invoke()
        assert questions[0]['default']=='no' and R.snap(r)==before
        with patch('stage8_closure.gui.messagebox.askyesno',return_value=True),patch('stage8_closure.gui.messagebox.showerror',side_effect=lambda *a:errors.append(a)):app.restore_button.invoke()
        assert not errors,errors
        success=app.text.get('1.0','end');assert 'RECOVERED_ABORTED' in success and 'Списаний: 0' in success
        assert not R.states(r).count('COMMITTED') and R.balances(r)['s8-a']==12
        window.destroy();window=tk.Tk();window.withdraw();app=App(window);app.load(R.req(r))
        assert str(app.close_button['state'])=='normal' and not app.confirm.get()
        return {'sandbox':str(r),'pending':text,'success':success,'default_no':True,'new_App_new_preview':True,'real_Tk_widgets':True,'visual_E2E':'NOT_RUN: native computer-control API unavailable'}
    finally:window.destroy()

def deep():
    r=OUT/'deep';path=r/('a'*110)/('b'*120)/'source.xlsm';path.parent.mkdir(parents=True)
    shutil.copyfile(RUN/'before_short.xlsm',path);before=S.sha(path)
    info=native.inspect(path,diagnostics=OUT/'office');assert info['sha256']==before and len(str(path))>350
    return {'source_length':len(str(path)),'short_length':len(info['verified_copy']['verification_path']),'native':info}

def native_save_death():
    import win32api,win32con,win32process
    r=R.fixture();p=R.pclose(r);R.save(r/'CLOSE_PREVIEW.json',p)
    w=R.proc('close_crash',r/'CLOSE_PREVIEW.json','before_save');assert w['exit']==73,w
    op=S.history(r)['active'];S.verify_backup(op['backup'],r)
    assert S.sha(R.reg(r))==op['before_sha256'] and R.states(r)==['INTENT','BACKUP_VERIFIED']
    records=[]
    for f in r.glob('transactions/*/office/*/START.json'):
        q=S.read(f)
        if not q['owned']:continue
        for name in q['expected_workbooks']:
            import re,win32security
            t=Path(name)
            # Windows packaged-app children may receive redirected LOCALAPPDATA.
            assert t.is_relative_to(Path.home()/'AppData/Local') and re.fullmatch('s8n_[a-f0-9]{32}',t.parent.name) and t.name=='stage.xlsm'
            sd=win32security.GetNamedSecurityInfo(str(t.parent),win32security.SE_FILE_OBJECT,win32security.OWNER_SECURITY_INFORMATION|win32security.DACL_SECURITY_INFORMATION)
            tok=win32security.OpenProcessToken(win32api.GetCurrentProcess(),win32con.TOKEN_QUERY)
            try:assert sd.GetSecurityDescriptorOwner()==win32security.GetTokenInformation(tok,win32security.TokenUser)[0]
            finally:tok.Close()
        if H.alive(q['pid']) and H.creation(q['pid'])==q['creation_time']:
            # Test-only orphan cleanup after intentional hard process death.
            # Exact PID+creation time from this child's ownership evidence.
            h=win32api.OpenProcess(win32con.PROCESS_QUERY_INFORMATION|win32con.PROCESS_TERMINATE,False,q['pid'])
            try:
                assert win32process.GetProcessTimes(h)['CreationTime'].isoformat()==q['creation_time']
                win32api.TerminateProcess(h,91)
            finally:h.Close()
            assert H.wait_exit(q['pid'])
        records.append(q)
    answer=R.result(R.proc('recover',r));assert answer['status']=='ABORTED_UNCHANGED'
    assert not R.states(r).count('COMMITTED') and R.balances(r)['s8-a']==12
    return {'sandbox':str(r),'death':w,'backup_intact':True,'recover':answer,'owned_test_orphan_records':records,'global_taskkill':False,'hardware_finally_guaranteed':False}

with office.diagnostic_scope(OUT/'office'):
    case('short_copy_external_writer_replace_native_parent_blocked',lambda:locks('copy'))
    case('original_source_external_writer_replace_native_parent_blocked',lambda:locks('source'))
    case('short_copy_changed_content_STOP_before_Excel',lambda:tamper('content'))
    case('short_copy_changed_identity_STOP_preserve_foreign',lambda:tamper('identity'))
    for mode in ('file','missing','long','acl'):case('close_preflight_'+mode+'_STOP_no_intent',lambda mode=mode:invalid(mode))
    case('restore_preflight_unavailable_STOP_no_rename',lambda:invalid('missing',True))
    case('resume_preflight_unavailable_STOP_no_rename',lambda:invalid('missing',True,True))
    case('verification_exception_cleans_private_workspace',exception_cleanup)
    case('GUI_long_path_real_widgets_restore_cancel_confirm_restart',gui)
    case('GUI_unavailable_workspace_safe_STOP',lambda:gui(True))
    case('native_source_path_over_350_same_bytes',deep)
    case('native_Save_process_death_backup_safe_no_COMMITTED',native_save_death)
    for name,fn in [('two_concurrent_closes_once',H.parallel),('changed_SHA_STOP',H.changed_sha),('same_SHA_changed_identity_STOP',H.identity_only),('existing_writer_STOP',H.preexisting_writer),('existing_writable_mapping_STOP',lambda:H.preexisting_writer(True)),('native_save_torn_private_copy_safe_STOP',H.torn_stage),('RPC_E_DISCONNECTED_exact_safe_STOP',H.disconnect),('real_COM_disconnect_at_Save_safe_STOP',lambda:H.disconnect(False)),('no_UZK_no_registry_write',H.no_uzk),('Workbook_Close_primary_hresult_preserved',H.book_close_primary)]:case(name,fn)
S.write(OUT/'PRODUCT_AFTER.json',{p:S.sha(p) for p in PRODUCT})
print('DONE',sum(q['status']=='PASS' for q in RESULTS),sum(q['status']=='FAIL' for q in RESULTS),flush=True)
