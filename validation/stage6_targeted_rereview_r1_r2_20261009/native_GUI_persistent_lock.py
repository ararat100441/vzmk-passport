from review_common import *
from akz_stage6 import core as c
from akz_stage6.gui import Launcher
import tkinter as tk
from tkinter import messagebox
import win32file,win32con
F=E/'sandbox/fixtures';O=E/'native_GUI_output';O.mkdir()
shutil.copytree(E/'sandbox/output/GOOD_N3',O/'OLD_GOOD')
oldsha={str(f):sha(f) for f in (O/'OLD_GOOD').iterdir()}
reg=F/'DEV_registry.xlsx';act=F/'ACT_template.xlsx';application=F/'APPLICATION_template.xlsx'
handles=[];proof=[];messages=[]
original=c.native_open;original_message=messagebox.showerror
def locked(excel,path,*,read_only):
    if Path(path).name==APPNAME and not read_only:
        healthy=sha(path);assert healthy==sha(application)
        handles.append(win32file.CreateFile(str(path),win32con.GENERIC_READ,0,None,win32con.OPEN_EXISTING,0,None))
        try:return original(excel,path,read_only=read_only)
        except Exception as exc:
            proof.append({'path':str(path),'healthy_sha_before':healthy,'share':0,'classified_code':getattr(exc,'code',None),'message':str(exc),'cause_winerror':getattr(exc.__cause__,'winerror',None)})
            raise
    return original(excel,path,read_only=read_only)
with patch.multiple(c,REGISTRY=reg,ACT=act,APPLICATION=application,FIXTURES=F,OUTPUT=O,PINS={act:sha(act),application:sha(application)}):
    root=tk.Tk();gui=Launcher(root)
    gui.vom.set(str(F/'selected_VOM.xls'));gui.catalog.set(str(F/'TEST_catalog.xlsx'))
    gui.check();assert gui.checked is not None and str(gui.generate_button['state'])=='normal'
    p=gui.checked
    def shown(title,message,**kw):
        messages.append({'title':title,'message':message})
        save('native_GUI_lock_message_before_OK.json',{'proof':proof,'messages':messages,'lock_still_held':bool(handles),'pending':[str(t) for t in O.glob('.pending-*')],'old_good_unchanged':all(sha(f)==h for f,h in oldsha.items())})
        print('REAL_GUI_MODAL_READY',flush=True)
        answer=original_message(title,message,**kw)
        root.after(100,root.quit)
        return answer
    print('REAL_GUI_READY_CLICK_GENERATE',flush=True)
    try:
        with patch.object(c,'native_open',side_effect=locked),patch('akz_stage6.gui.messagebox.showerror',side_effect=shown):root.mainloop()
        pending=list(O.glob('.pending-*'))
        save('native_GUI_lock_after_OK.json',{'proof':proof,'messages':messages,'pending':[str(t) for t in pending],'preview_invalidated':gui.checked is None,'published':[str(t) for t in O.iterdir() if not t.name.startswith('.pending-') and t.name!='OLD_GOOD'],'old_good_unchanged':all(sha(f)==h for f,h in oldsha.items())})
    finally:
        for h in handles:h.Close()
        root.destroy()
    retry=c.generate(p,'RETRY_RELEASED')
    assert load(retry/'PAIR.json')['state']=='COMPLETE_TEST_PAIR'
    for t in list(O.glob('.pending-*')):
        assert t.resolve().parent==O.resolve() and O.resolve().is_relative_to(E.resolve()) and t.name.startswith('.pending-')
        shutil.rmtree(t)
    save('native_GUI_release_retry.json',{'target':str(retry),'sha':{str(f):sha(f) for f in retry.iterdir()},'old_good_unchanged':all(sha(f)==h for f,h in oldsha.items()),'pending_removed_after_capture':[str(t) for t in pending],'audit':AUDIT})
    print('REAL_GUI_RETRY_COMPLETE',flush=True)
