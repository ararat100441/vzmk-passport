"""Source experiment comparing registered COM startup with the old mixed flags."""
from common import *
import time,gc,win32process,win32api,win32event,pythoncom,winreg,traceback,win32gui
from stage9_app.ownership import office_image,process_image,native_application
from openpyxl import Workbook
RUN=OUT/('launch_protocol_'+uuid4().hex[:8]);RUN.mkdir();cases=[]
safe=len(sys.argv)>1 and sys.argv[1]=='safe'
window_close=len(sys.argv)>1 and sys.argv[1]=='window-close'
def registry_state():
    result={}
    for hive,name in ((winreg.HKEY_LOCAL_MACHINE,'HKLM'),(winreg.HKEY_CURRENT_USER,'HKCU')):
        for appname in ('Excel','Word'):
            key='SOFTWARE\\Microsoft\\Office\\'+appname+'\\Addins\\PDFMaker.OfficeAddin'
            try:
                with winreg.OpenKey(hive,key) as k:result[name+'/'+key]=winreg.QueryValueEx(k,'LoadBehavior')[0]
            except FileNotFoundError:result[name+'/'+key]='ABSENT'
    return result
before=registry_state();save(RUN/'ADDIN_CONFIGURATION_BEFORE.json',before)
for kind,flags in ([('excel',['/safe'])] if safe else [('excel',['/automation','-Embedding'])] if window_close else [('excel',['/automation','-Embedding']),('word',['/automation','-Embedding'])]):
    progid='Excel.Application' if kind=='excel' else 'Word.Application';path=RUN/('probe.xlsx' if kind=='excel' else 'probe.docx')
    if kind=='excel':
        b=Workbook();b.save(path);b.close()
    else:
        from docx import Document
        Document().save(path)
    startup=subprocess.STARTUPINFO();startup.dwFlags|=subprocess.STARTF_USESHOWWINDOW;startup.wShowWindow=0
    child=subprocess.Popen([office_image(progid),*flags,str(path)],startupinfo=startup,creationflags=subprocess.CREATE_NO_WINDOW)
    record={'pid':child.pid,'creation_time':win32process.GetProcessTimes(child._handle)['CreationTime'].isoformat(),'image':process_image(child._handle),'hwnd':0,'worker_pid':os.getpid(),'launcher_pid':os.getpid(),'actual_child':True,'handle_source':'CREATE_PROCESS'}
    save(RUN/(kind+'_CREATED.json'),record|{'arguments':flags,'probe':str(path)})
    pythoncom.CoInitialize();app=None;started=time.time();error=None;info={}
    try:
        deadline=time.monotonic()+30
        while True:
            try:
                with native_application(record,child._handle) as native:app=native
                native=None;break
            except RuntimeError as exc:
                if str(exc)!='OWNED_NATIVEOM_WINDOW_MISSING' or time.monotonic()>=deadline:raise
                time.sleep(.1)
        app.AutomationSecurity=3;app.DisplayAlerts=0 if kind=='word' else False
        info={'user_control':bool(app.UserControl),'version':str(app.Version)}
        if kind=='excel':
            info['workbooks']=int(app.Workbooks.Count)
            try:info['connected_com_addins']=[a.ProgId for a in app.COMAddIns if a.Connect]
            except Exception as exc:info['com_addins_query_error']=repr(exc)
            info['pdfmaker_modules']=[win32process.GetModuleFileNameEx(child._handle,m) for m in win32process.EnumProcessModules(child._handle) if 'pdfm' in win32process.GetModuleFileNameEx(child._handle,m).casefold()]
            app.UserControl=False
            while app.Workbooks.Count:app.Workbooks(1).Close(False)
            app.Quit()
        else:app.Documents.Close(False);app.Quit(SaveChanges=0)
    except Exception as exc:error=repr(exc)
    finally:
        if error and app is not None:
            try:app.Quit()
            except Exception:pass
        app=None;gc.collect();info['interfaces_after_gc']=pythoncom._GetInterfaceCount()
        deadline=time.monotonic()+15
        while child.poll() is None and time.monotonic()<deadline:pythoncom.PumpWaitingMessages();time.sleep(.1)
        if child.poll() is None and window_close:
            assert win32process.GetProcessTimes(child._handle)['CreationTime'].isoformat()==record['creation_time'] and process_image(child._handle)==record['image']
            windows=[]
            win32gui.EnumWindows(lambda hwnd,_:windows.append(hwnd) if win32process.GetWindowThreadProcessId(hwnd)[1]==record['pid'] and win32gui.GetClassName(hwnd)=='XLMAIN' else None,None)
            save(RUN/'NATIVE_WM_CLOSE_REQUEST.json',record|{'windows':windows,'original_handle':int(child._handle),'workbooks_closed':True,'quit_returned':error is None,'interfaces_after_gc':info['interfaces_after_gc']})
            for hwnd in windows:
                if win32process.GetWindowThreadProcessId(hwnd)[1]==record['pid']:win32gui.PostMessage(hwnd,0x10,0,0)
            deadline=time.monotonic()+15
            while child.poll() is None and time.monotonic()<deadline:pythoncom.PumpWaitingMessages();time.sleep(.1)
        pythoncom.CoUninitialize()
        if child.poll() is None:
            # Emergency only for this actual child on its retained original handle.
            assert win32process.GetProcessTimes(child._handle)['CreationTime'].isoformat()==record['creation_time'] and process_image(child._handle)==record['image'] and win32process.GetProcessId(child._handle)==record['pid']
            save(RUN/(kind+'_FORCE_BEGIN.json'),record|{'original_handle':int(child._handle),'identity_validated':True,'reason':error or 'SOURCE_EXPERIMENT_GRACE_TIMEOUT'})
            win32api.TerminateProcess(child._handle,1);child.wait(timeout=10)
    cases.append({'kind':kind,'source_experiment_only':True,'info':info,'error':error,'exit_code':child.returncode,'seconds':time.time()-started})
    print(kind,cases[-1],flush=True)
save(RUN/'RESULT.json',{'source_experiment_only':True,'cases':cases})
after=registry_state();save(RUN/'ADDIN_CONFIGURATION_AFTER.json',after);assert before==after,(before,after)
