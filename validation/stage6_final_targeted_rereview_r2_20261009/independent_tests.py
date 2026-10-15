from common_review import *
from akz_stage6 import core as c
from akz_stage6.gui import Launcher
from openpyxl import load_workbook
import tkinter as tk
import win32file,win32con,pywintypes,subprocess,unittest,io
from contextlib import nullcontext
F,O=setup('sandbox');sources=snapshot(F);oldgood=snapshot(O/'OLD_GOOD');userfile=sha(O/'user_document.xlsx')
results=[]
def check(name,fn):
    try:
        detail=fn();results.append({'name':name,'status':'PASS','detail':detail});print('PASS',name,flush=True)
    except Exception as exc:
        results.append({'name':name,'status':'FAIL','error':str(exc),'traceback':traceback.format_exc()});print('FAIL',name,str(exc),flush=True);raise
    finally:save('independent_tests.json',{'results':results,'audit':AUDIT})
def yes(v):assert v
def stat_snapshot(p):return {f.name:[f.lstat().st_dev,f.lstat().st_ino,f.lstat().st_size,f.lstat().st_mtime_ns] for f in p.iterdir()}
def saved_pair(t,p):
    r=load(t/'PAIR.json');assert r['state']=='COMPLETE_TEST_PAIR' and r['number']==p.payload['act_number'] and r['date']==p.payload['date']
    assert r['sha256']=={n:sha(t/n) for n in (ACTNAME,APPNAME)}
    assert set(x.name for x in t.iterdir())=={ACTNAME,APPNAME,'PAIR.json'}
    return {'target':str(t),'receipt':r}
def smoke(t,p):
    with c.native_excel() as app:
        b=c.native_open(app,t/APPNAME,read_only=True)
        try:
            s=b.Worksheets(1);n=len(p.payload['selected'])
            hidden=[r for r in range(1,22) if s.Rows(r).Hidden]
            assert hidden==list(range(n+4,21)) and not s.Rows(21).Hidden
            totals={col+'21':s.Range(col+'21').Value2 for col in 'CEG'}
            for col,key in zip('CEG',('quantity','mass_total','area_total')):
                assert s.Range(col+'21').Formula==f'=SUM({col}4:{col}20)'
                assert abs(Decimal(str(totals[col+'21']))-p.payload['totals'][key])<Decimal('1e-8')
            assert [s.Range('A'+str(r)).Value2 for r in range(4,n+4)]==[row['mark'] for row in p.payload['selected']]
            return {'read_only':bool(b.ReadOnly),'N':n,'hidden':hidden,'total_visible':True,'totals':totals}
        finally:b.Close(SaveChanges=False)
with isolated(c,F,O):
    p=c.preview('Гараж Казахстан','9/2026',F/'selected_VOM.xls',F/'TEST_catalog.xlsx')
    save('independent_preview.json',{'seal':p.seal,'payload':p.payload})
    assert p.payload['totals']=={'quantity':Decimal(3),'mass_total':Decimal('105.26'),'area_total':Decimal('3.989')}
    handles=[];primary=[];proof=[];original_open=c.native_open;seen=[];original_error=c.error_message
    def locked(app,path,*,read_only):
        if Path(path).name==APPNAME and not read_only:
            healthy=sha(path);assert healthy==sha(F/'APPLICATION_template.xlsx')
            h=win32file.CreateFile(str(path),win32con.GENERIC_READ,0,None,win32con.OPEN_EXISTING,0,None);handles.append(h)
            try:return original_open(app,path,read_only=read_only)
            except Exception as exc:
                primary.append(exc);proof.append({'path':str(path),'healthy_sha':healthy,'share':0,'code':getattr(exc,'code',None),'winerror':getattr(exc.__cause__,'winerror',None)});raise
        return original_open(app,path,read_only=read_only)
    def error(exc):seen.append(exc);return original_error(exc)
    root=tk.Tk();root.withdraw();gui=Launcher(root);gui.checked=p
    try:
        with patch.object(c,'native_open',side_effect=locked),patch.object(c,'error_message',side_effect=error),patch('akz_stage6.gui.messagebox.showerror') as msg:gui.generate()
        message=msg.call_args.args[1];temp=Path(proof[0]['path']).parent;ledger=next(O.glob('.stage6-pending-*.json'));record=load(ledger)
        cleanup=getattr(primary[0],'cleanup_error',None)
        def preserve():
            assert seen[0] is primary[0] and primary[0].code=='FILE_LOCKED' and cleanup.code=='FILE_LOCKED'
            assert primary[0].__cause__.winerror==32 and cleanup.__cause__.winerror==32
            assert message.startswith('FILE_LOCKED:') and proof[0]['path'] in message and 'Файл приложения АКЗ занят или заблокирован' in message and 'Закройте файл в Excel и повторите операцию' in message
            assert 'OTHER_ERROR' not in message and 'INVALID_XLSX' not in message and 'WinError 32' not in message
            assert all(win32file.GetFileType(h)==win32con.FILE_TYPE_DISK for h in handles) and gui.checked is None
            detail={'proof':proof,'GUI_message':message,'primary_same_object_at_GUI':True,'cleanup_diagnostic':{'code':cleanup.code,'message':str(cleanup),'cause_winerror':cleanup.__cause__.winerror},'lock_live_after_GUI_return':True}
            save('real_lock_primary_proof.json',detail);return detail
        check('real_lock_primary_same_object_through_GUI_return',preserve)
        def ownership():
            import uuid
            assert record['id']==uuid.UUID(record['id']).hex and temp==O.resolve()/('.pending-stage6-'+record['id'])
            assert record['root']==str(O.resolve()) and record['path']==str(temp)
            assert record['directory_identity']==[temp.lstat().st_dev,temp.lstat().st_ino]
            assert set(record['files'])=={ACTNAME,APPNAME}
            assert all(record['files'][f.name]==[f.lstat().st_dev,f.lstat().st_ino] for f in temp.iterdir())
            assert set(x.name for x in O.iterdir() if not x.name.startswith('.'))=={'OLD_GOOD','user_document.xlsx'}
            return record
        check('pending_UUID_exact_path_and_Windows_identity',ownership)
        lockedstats=stat_snapshot(temp);ledgerhash=sha(ledger)
        def retry_locked():
            r=stop(lambda:c.generate(p,'STILL_LOCKED'),'FILE_LOCKED')
            assert stat_snapshot(temp)==lockedstats and sha(ledger)==ledgerhash and not (O/'STILL_LOCKED').exists();return r
        check('retry_while_locked_preserves_all_files_and_ledger',retry_locked)
    finally:
        for h in handles:h.Close()
        root.destroy()
    private=snapshot(temp)
    unknown=O/'.pending-other-owner';unknown.mkdir();(unknown/'USER.xlsx').write_bytes(b'FOREIGN OWNER')
    before=snapshot(O)
    def unknown_test():
        r=stop(c.recover_pending,'PENDING_OWNERSHIP_STOP');assert snapshot(O)==before;return r
    check('unknown_pending_STOP_before_any_owned_deletion',unknown_test)
    (unknown/'USER.xlsx').unlink();unknown.rmdir()
    foreign=temp/'USER.xlsx';foreign.write_bytes(b'UNKNOWN FILE INSIDE OWNED PENDING');before=snapshot(O)
    check('unknown_content_inside_owned_pending_STOP',unknown_test)
    foreign.unlink()
    application=temp/APPNAME;held=E/'held_original_application.xlsx';application.rename(held);shutil.copyfile(held,application)
    before=snapshot(O);assert application.lstat().st_ino!=record['files'][APPNAME][1]
    check('replacement_file_identity_STOP_no_deletion',unknown_test)
    application.unlink();held.rename(application)
    held_dir=E/'held_original_pending';temp.rename(held_dir);temp.mkdir();shutil.copyfile(held_dir/ACTNAME,temp/ACTNAME);shutil.copyfile(held_dir/APPNAME,temp/APPNAME)
    before=snapshot(O);assert temp.lstat().st_ino!=record['directory_identity'][1]
    check('replacement_directory_identity_STOP_no_deletion',unknown_test)
    # Remove only this harness-created replacement directory, then restore real identities.
    assert temp.resolve().parent==O.resolve() and O.resolve().is_relative_to(E.resolve())
    for f in list(temp.iterdir()):f.unlink()
    temp.rmdir();held_dir.rename(temp)
    ledgerbytes=ledger.read_bytes();tampered=copy.deepcopy(record);tampered['path']=str(O/'OLD_GOOD');ledger.write_text(json.dumps(tampered),encoding='utf-8');before=snapshot(O)
    check('tampered_record_path_STOP_no_deletion',unknown_test);ledger.write_bytes(ledgerbytes)
    tampered=copy.deepcopy(record);tampered['id']='NOT A UUID';ledger.write_text(json.dumps(tampered),encoding='utf-8');before=snapshot(O)
    check('invalid_UUID_record_STOP_no_deletion',unknown_test);ledger.write_bytes(ledgerbytes)
    assert snapshot(temp)==private
    def child_recovery():
        r=subprocess.run([sys.executable,'-B',str(E/'recover_child.py'),str(O)],capture_output=True,text=True,encoding='utf-8',timeout=45)
        assert r.returncode==0,(r.stdout,r.stderr)
        assert not temp.exists() and not ledger.exists() and snapshot(O/'OLD_GOOD')==oldgood and sha(O/'user_document.xlsx')==userfile
        return {'child_pid_independent':True,'stdout':r.stdout,'stderr':r.stderr,'record':record,'owned_pending_removed':True,'foreign_files_unchanged':True}
    check('fresh_Python_process_recovers_owned_pending_after_release',child_recovery)
    target=c.generate(p,'RETRY_AFTER_RELEASE')
    check('retry_complete_pair_after_FILE_LOCKED',lambda:saved_pair(target,p))
    check('R1_N3_native_smoke',lambda:smoke(target,p))
    p17=c.preview('Гараж Казахстан','9/2026',F/'SYNTHETIC_17.xlsx',F/'TEST_catalog.xlsx',synthetic=True);t17=c.generate(p17,'N17')
    check('R1_N17_native_smoke',lambda:smoke(t17,p17))
    def capacity():
        before=snapshot(O);r=stop(lambda:c.preview('Гараж Казахстан','9/2026',F/'SYNTHETIC_18.xlsx',F/'TEST_catalog.xlsx',synthetic=True),'TEMPLATE_CAPACITY_STOP');assert snapshot(O)==before;return r
    check('R1_N18_CAPACITY_STOP_no_publication',capacity)
    def corrupt(app,path,*,read_only):
        if Path(path).name==APPNAME and not read_only:Path(path).write_bytes(b'CORRUPT ISOLATED XLSX')
        return original_open(app,path,read_only=read_only)
    def corrupt_test():
        before=snapshot(O)
        with patch.object(c,'native_open',side_effect=corrupt):r=stop(lambda:c.generate(p,'CORRUPT'),'INVALID_XLSX')
        assert snapshot(O)==before;return r
    check('actual_corrupt_XLSX_INVALID_XLSX_cleanup_complete',corrupt_test)
    def exists():
        before=snapshot(target);r=stop(lambda:c.generate(p,'RETRY_AFTER_RELEASE'),'OUTPUT_ALREADY_EXISTS_STOP');assert snapshot(target)==before;return r
    check('OUTPUT_EXISTS_preserves_successful_pair',exists)
    for path in (c.REGISTRY,c.ACT,c.APPLICATION,F/'TEST_catalog.xlsx',F/'selected_VOM.xls'):
        data=path.read_bytes()
        try:
            path.write_bytes(data+b'ISOLATED INPUT CHANGE')
            check('INPUT_CHANGED_'+path.name,lambda:stop(lambda:c.generate(p,'CHANGED'),'INPUT_CHANGED_AFTER_PREVIEW_STOP'))
        finally:path.write_bytes(data)
    with patch('win32file.CreateFile',side_effect=pywintypes.error(5,'CreateFile','Access denied')):
        check('ACCESS_DENIED_unit_error5',lambda:stop(lambda:c.probe_file_access(c.APPLICATION,read_only=False),'ACCESS_DENIED'))
    check('OTHER_ERROR_only_generic_valid_XLSX_failure',lambda:stop(lambda:c.excel_file_failure(c.APPLICATION,RuntimeError('generic non-format failure'),read_only=True,operation='review'),'OTHER_ERROR'))
    for code in ('FILE_LOCKED','ACCESS_DENIED','INVALID_XLSX','OTHER_ERROR','OUTPUT_ALREADY_EXISTS_STOP','INPUT_CHANGED_AFTER_PREVIEW_STOP'):
        primary_error=c.FileOperationError(code,c.APPLICATION,'primary proof') if code in ('FILE_LOCKED','ACCESS_DENIED','INVALID_XLSX','OTHER_ERROR') else c.Stage6Error(code)
        secondary=PermissionError(13,'independent secondary failure');secondary.winerror=32
        def primary_test():
            with patch.object(c,'native_excel',side_effect=lambda:nullcontext(None)),patch.object(c,'write_act',side_effect=primary_error),patch.object(c,'cleanup_pending',side_effect=secondary):
                try:c.generate(p,'PRIMARY_'+code)
                except Exception as exc:assert exc is primary_error and exc.cleanup_error is secondary
                else:raise AssertionError('PRIMARY LOST')
            c.recover_pending();return {'code':code,'same_primary_object':True,'secondary_diagnostic_only':True}
        check('primary_preserved_on_secondary_cleanup_'+code,primary_test)
    check('all_isolated_sources_old_pair_user_file_preserved',lambda:yes(snapshot(F)==sources and snapshot(O/'OLD_GOOD')==oldgood and sha(O/'user_document.xlsx')==userfile and not list(O.glob('.pending-*')) and not list(O.glob('.stage6-pending-*.json'))))
suite=unittest.defaultTestLoader.discover(str(ROOT/'tests'),pattern='test_stage6.py');log=io.StringIO();unit=unittest.TextTestRunner(stream=log,verbosity=2).run(suite)
(E/'unit_tests_independent.txt').write_text(log.getvalue(),encoding='utf-8');assert unit.wasSuccessful() and unit.testsRun==16
save('independent_tests.json',{'results':results,'summary':{'PASS':sum(r['status']=='PASS' for r in results),'FAIL':sum(r['status']=='FAIL' for r in results)},'units':{'run':unit.testsRun,'pass':unit.wasSuccessful()},'audit':AUDIT})
print('INDEPENDENT',len(results),'PASS, UNIT16 PASS',flush=True)
