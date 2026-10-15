from common_review import *
from akz_stage6 import core as c
from akz_stage6.gui import Launcher
import tkinter as tk
from tkinter import messagebox
import win32file,win32con
F,O=setup('native_GUI');oldgood=snapshot(O/'OLD_GOOD');unrelated=sha(O/'user_document.xlsx')
product=load(E/'baseline_current.json')['product']['CURRENT'];assert {p:sha(p) for p in product}==product
handles=[];primary=[];proof=[];seen=[];messages=[]
original_open=c.native_open;original_error=c.error_message;original_show=messagebox.showerror
def locked(app,path,*,read_only):
    if Path(path).name==APPNAME and not read_only:
        healthy=sha(path);assert healthy==sha(F/'APPLICATION_template.xlsx')
        handles.append(win32file.CreateFile(str(path),win32con.GENERIC_READ,0,None,win32con.OPEN_EXISTING,0,None))
        try:return original_open(app,path,read_only=read_only)
        except Exception as exc:
            primary.append(exc);proof.append({'path':str(path),'healthy_sha_before_lock':healthy,'share':0,'primary_code':getattr(exc,'code',None),'primary_cause_winerror':getattr(exc.__cause__,'winerror',None)});raise
    return original_open(app,path,read_only=read_only)
def error(exc):seen.append(exc);return original_error(exc)
with isolated(c,F,O):
    root=tk.Tk();gui=Launcher(root)
    gui.vom.set(str(F/'selected_VOM.xls'));gui.catalog.set(str(F/'TEST_catalog.xlsx'));gui.check();p=gui.checked
    assert p is not None and str(gui.generate_button['state'])=='normal'
    def shown(title,message,**kw):
        assert primary[0] is seen[0] and primary[0].code=='FILE_LOCKED' and primary[0].cleanup_error.code=='FILE_LOCKED'
        assert primary[0].__cause__.winerror==32 and primary[0].cleanup_error.__cause__.winerror==32
        assert message.startswith('FILE_LOCKED') and 'Закройте файл в Excel и повторите операцию' in message and proof[0]['path'] in message
        assert 'OTHER_ERROR' not in message and 'INVALID_XLSX' not in message and 'WinError 32' not in message
        assert all(win32file.GetFileType(h)==win32con.FILE_TYPE_DISK for h in handles)
        messages.append({'title':title,'message':message})
        save('native_GUI_modal_lock_held.json',{'proof':proof,'messages':messages,'same_primary_object':True,'cleanup_diagnostic':{'code':primary[0].cleanup_error.code,'cause_winerror':primary[0].cleanup_error.__cause__.winerror,'message':str(primary[0].cleanup_error)},'lock_live_during_actual_modal':True,'ledger':[load(x) for x in O.glob('.stage6-pending-*.json')]})
        print('ACTUAL_MODAL_READY_LOCK_HELD',flush=True)
        answer=original_show(title,message,**kw);root.after(100,root.quit);return answer
    print('REAL_GUI_READY_CLICK_GENERATE',flush=True)
    try:
        with patch.object(c,'native_open',side_effect=locked),patch.object(c,'error_message',side_effect=error),patch('akz_stage6.gui.messagebox.showerror',side_effect=shown):root.mainloop()
        assert handles and all(win32file.GetFileType(h)==win32con.FILE_TYPE_DISK for h in handles)
        pending=list(O.glob('.pending-*'));ledgers=list(O.glob('.stage6-pending-*.json'))
        assert len(pending)==len(ledgers)==1
        record=load(ledgers[0]);temp=pending[0]
        assert record['directory_identity']==[temp.lstat().st_dev,temp.lstat().st_ino]
        assert all(record['files'][f.name]==[f.lstat().st_dev,f.lstat().st_ino] for f in temp.iterdir())
        assert set(x.name for x in O.iterdir() if not x.name.startswith('.'))=={'OLD_GOOD','user_document.xlsx'}
        save('native_GUI_after_return_lock_held.json',{'GUI_handler_returned':True,'lock_still_live':True,'preview_invalidated':gui.checked is None,'owned_record':record,'pending_files':[str(x) for x in temp.iterdir()],'old_good_unchanged':snapshot(O/'OLD_GOOD')==oldgood,'no_partial_publication':True})
    finally:
        for h in handles:h.Close()
        root.destroy()
    target=c.generate(p,'NORMAL_RETRY_AFTER_RELEASE')
    assert load(target/'PAIR.json')['state']=='COMPLETE_TEST_PAIR'
    assert not temp.exists() and not ledgers[0].exists() and not list(O.glob('.pending-*')) and not list(O.glob('.stage6-pending-*.json'))
    assert snapshot(O/'OLD_GOOD')==oldgood and sha(O/'user_document.xlsx')==unrelated
    save('native_GUI_release_product_cleanup_retry.json',{'normal_generate_removed_own_old_pending':True,'no_harness_cleanup':True,'complete_pair':str(target),'receipt':load(target/'PAIR.json'),'pair_files':snapshot(target),'old_good_unchanged':True,'unrelated_unchanged':True,'audit':AUDIT})
    assert {p:sha(p) for p in product}==product
    print('REAL_GUI_RETRY_AND_PRODUCT_CLEANUP_PASS',flush=True)
