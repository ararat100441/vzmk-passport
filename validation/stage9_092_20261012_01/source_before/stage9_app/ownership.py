"""Bind DispatchEx to actual Office windows; graceful teardown, no global kill."""
from contextlib import contextmanager
from pathlib import Path
import os
import time

def process_image(handle):
    import ctypes
    from ctypes import wintypes as W
    query=ctypes.WinDLL('kernel32',use_last_error=True).QueryFullProcessImageNameW
    query.argtypes=[W.HANDLE,W.DWORD,W.LPWSTR,ctypes.POINTER(W.DWORD)];query.restype=W.BOOL
    buffer=ctypes.create_unicode_buffer(32768);length=W.DWORD(len(buffer))
    if not query(int(handle),0,buffer,ctypes.byref(length)):raise ctypes.WinError(ctypes.get_last_error())
    return buffer.value

def office_image(progid):
    import winreg
    name='EXCEL.EXE' if str(progid).casefold()=='excel.application' else 'WINWORD.EXE'
    for hive in (winreg.HKEY_LOCAL_MACHINE,winreg.HKEY_CURRENT_USER):
        try:
            with winreg.OpenKey(hive,'SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\App Paths\\'+name) as key:path=Path(winreg.QueryValueEx(key,None)[0])
            if path.is_file() and path.name.upper()==name:return str(path.resolve())
        except OSError:pass
    raise RuntimeError('OFFICE_EXECUTABLE_NOT_REGISTERED')

def launch_child(progid,probe,worker_pid):
    """The caller keeps the original CreateProcess handle, never a DCOM PID guess."""
    import subprocess,ctypes,win32process
    from ctypes import wintypes as W
    image=office_image(progid);word=str(progid).casefold()=='word.application'
    switches=['/w','/m'] if word else ['/x','/automation']
    startup=subprocess.STARTUPINFO();startup.dwFlags|=subprocess.STARTF_USESHOWWINDOW;startup.wShowWindow=0
    kernel=ctypes.WinDLL('kernel32',use_last_error=True)
    get=kernel.GetDllDirectoryW;get.argtypes=[W.DWORD,W.LPWSTR];get.restype=W.DWORD
    setdir=kernel.SetDllDirectoryW;setdir.argtypes=[W.LPCWSTR];setdir.restype=W.BOOL
    buffer=ctypes.create_unicode_buffer(32768);get(len(buffer),buffer)
    if not setdir(None):raise ctypes.WinError(ctypes.get_last_error())
    try:child=subprocess.Popen([image,*switches,str(probe)],startupinfo=startup,creationflags=subprocess.CREATE_NO_WINDOW)
    finally:setdir(buffer.value or None)
    record={'pid':child.pid,'creation_time':win32process.GetProcessTimes(child._handle)['CreationTime'].isoformat(),
            'image':process_image(child._handle),'hwnd':0,'progid':progid,'worker_pid':worker_pid,
            'launcher_pid':os.getpid(),'actual_child':True,'owned':True,'absent_before_dispatch':True,'quit_attempted':False,
            'retained_handle':int(child._handle),'handle_source':'CREATE_PROCESS'}
    if record['image'].casefold()!=image.casefold():raise RuntimeError('OFFICE_CHILD_IMAGE_CHANGED')
    return child,record

@contextmanager
def native_application(record,handle,trace=lambda *args:None):
    """Bind only a window in the PID whose original process handle is retained."""
    import ctypes,uuid,gc,pythoncom,win32com.client,win32gui,win32process
    from ctypes import wintypes as W
    if win32process.GetProcessTimes(handle)['CreationTime'].isoformat()!=record['creation_time'] or process_image(handle)!=record['image']:raise RuntimeError('OFFICE_IDENTITY_MISMATCH_NO_QUIT')
    candidates=[];expected='_WwG' if Path(record['image']).name.upper()=='WINWORD.EXE' else 'EXCEL7'
    def window(hwnd,_):
        if win32process.GetWindowThreadProcessId(hwnd)[1]==record['pid']:
            priority=(hwnd!=record['hwnd'],not bool(win32gui.GetWindowText(hwnd)))
            def child(child_hwnd,_):
                if win32gui.GetClassName(child_hwnd)==expected and win32process.GetWindowThreadProcessId(child_hwnd)[1]==record['pid']:candidates.append((priority,child_hwnd))
            win32gui.EnumChildWindows(hwnd,child,None)
    win32gui.EnumWindows(window,None)
    if not candidates:raise RuntimeError('OWNED_NATIVEOM_WINDOW_MISSING')
    obj=native=app=None
    try:
        fn=ctypes.OleDLL('oleacc').AccessibleObjectFromWindow
        fn.argtypes=[W.HWND,W.DWORD,ctypes.c_void_p,ctypes.POINTER(ctypes.c_void_p)];fn.restype=ctypes.c_long
        iid=ctypes.create_string_buffer(uuid.UUID('00020400-0000-0000-c000-000000000046').bytes_le);pointer=ctypes.c_void_p()
        trace('CANDIDATES',candidates)
        for _,hwnd in sorted(candidates):
            try:
                trace('ACCESSIBLE_BEGIN',hwnd);hr=fn(hwnd,0xfffffff0,ctypes.byref(iid),ctypes.byref(pointer))
                if hr==0:break
            except OSError as exc:trace('ACCESSIBLE_ERROR',repr(exc))
        else:raise RuntimeError('OWNED_NATIVEOM_UNAVAILABLE_NO_QUIT')
        trace('ACCESSIBLE_RETURNED',hwnd)
        try:obj=pythoncom.ObjectFromAddress(pointer.value,pythoncom.IID_IDispatch)
        finally:
            # ObjectFromAddress adds its own QI reference; release the original.
            table=ctypes.cast(pointer,ctypes.POINTER(ctypes.POINTER(ctypes.c_void_p))).contents
            ctypes.WINFUNCTYPE(W.ULONG,ctypes.c_void_p)(table[2])(pointer)
        native=win32com.client.dynamic.Dispatch(obj)
        if win32process.GetWindowThreadProcessId(native.Hwnd)[1]!=record['pid']:raise RuntimeError('NATIVEOM_PID_MISMATCH_NO_QUIT')
        app=native.Application
        app_pid=win32process.GetWindowThreadProcessId(app.Hwnd if expected=='EXCEL7' else native.Hwnd)[1]
        if app_pid!=record['pid']:raise RuntimeError('APPLICATION_PID_MISMATCH_NO_QUIT')
        trace('OWN_APPLICATION_BOUND',record['pid'])
        yield app
    finally:app=native=obj=None;gc.collect()

def quit_owned_window(record,handle,trace=lambda *args:None):
    import pythoncom
    pythoncom.CoInitialize()
    try:
        with native_application(record,handle,trace) as app:
            if Path(record['image']).name.upper()=='WINWORD.EXE':app.Quit(SaveChanges=0)
            else:
                app.DisplayAlerts=False
                while app.Workbooks.Count:app.Workbooks(1).Close(SaveChanges=False)
                app.Quit()
            trace('QUIT_RETURNED',record['pid'])
        app=None
        return {'status':'PASS','action':'NATIVEOM_CLOSE_WITHOUT_SAVE_AND_QUIT','pid':record['pid'],'creation_time':record['creation_time']}
    finally:pythoncom.CoUninitialize()

class OwnedApplication:
    def __init__(self,app,record,event):
        object.__setattr__(self,'_app',app);object.__setattr__(self,'_record',record);object.__setattr__(self,'_event',event)
    def __getattr__(self,name):return getattr(self._app,name)
    def __setattr__(self,name,value):setattr(self._app,name,value)
    def Quit(self,*args,**kwargs):
        self._record['quit_attempted']=True;self._event('QUIT_BEGIN',self._record)
        try:
            value=self._app.Quit(*args,**kwargs)
            self._record['quit_returned']=True;self._event('QUIT_RETURNED',self._record)
            return value
        except BaseException as exc:
            self._record['quit_error']=repr(exc);self._event('QUIT_ERROR',self._record);raise

@contextmanager
def owned_office(diagnostic, *, supervised=False):
    import win32com.client,win32process,win32api,win32con,win32event,pythoncom
    from .workspace import save
    from uuid import uuid4
    diagnostic=Path(diagnostic)
    def event(kind,data):
        save(diagnostic.with_name(diagnostic.stem+'_'+kind+'_'+uuid4().hex[:8]+'.json'),
             {'event':kind,'monotonic':time.monotonic(),'worker_pid':os.getpid(),'data':dict(data)})
    original=win32com.client.DispatchEx;pythoncom.CoInitialize()
    uninitialize=pythoncom.CoUninitialize;deferred=[]
    pythoncom.CoUninitialize=lambda:deferred.append(True)
    owned=[];records=[];processes=[]
    def dispatch(progid,*args,**kwargs):
        kind=str(progid).casefold()
        if kind not in ('excel.application','word.application'):return original(progid,*args,**kwargs)
        suffix='.docx' if kind=='word.application' else '.xlsx'
        probe=diagnostic.parent/('probe_'+uuid4().hex[:8]+suffix)
        if suffix=='.xlsx':
            from openpyxl import Workbook
            book=Workbook();book.save(probe);book.close()
        else:
            from docx import Document
            Document().save(probe)
        if supervised:
            from stage8_closure import storage as S
            request=diagnostic.parent/('launch_'+uuid4().hex+'.json')
            save(request,{'progid':progid,'probe':str(probe),'probe_sha256':S.sha(probe),'worker_pid':os.getpid()})
            response=request.with_suffix('.response.json');deadline=time.monotonic()+30
            while not response.exists() and time.monotonic()<deadline:time.sleep(.05)
            if not response.exists():raise RuntimeError('OFFICE_LAUNCH_RESPONSE_TIMEOUT')
            reply=S.read(response)
            if reply['status']!='PASS':raise RuntimeError('OFFICE_LAUNCH_FAILED: '+reply['error'])
            record=reply['record'];handle=win32api.OpenProcess(win32con.PROCESS_QUERY_INFORMATION|win32con.SYNCHRONIZE,False,record['pid'])
        else:
            child,record=launch_child(progid,probe,os.getpid());processes.append(child);handle=child._handle
        owned.append(handle);records.append(record)
        deadline=time.monotonic()+30
        while True:
            try:
                with native_application(record,handle) as native:app=native
                native=None;break
            except RuntimeError as exc:
                if str(exc)!='OWNED_NATIVEOM_WINDOW_MISSING' or time.monotonic()>=deadline:raise
                time.sleep(.1)
        app.AutomationSecurity=3;app.DisplayAlerts=0 if suffix=='.docx' else False;app.Visible=False
        if suffix=='.docx':
            record['hwnd']=int(app.Windows(1).Hwnd);app.Documents.Close(False)
        else:
            record['hwnd']=int(app.Hwnd)
            while app.Workbooks.Count:app.Workbooks(1).Close(SaveChanges=False)
        save(diagnostic.with_name(diagnostic.stem+f'_{record["pid"]}_START.json'),dict(record))
        return OwnedApplication(app,record,event)
    win32com.client.DispatchEx=dispatch
    try:yield records
    finally:
        win32com.client.DispatchEx=original
        import gc;gc.collect()
        pythoncom.CoUninitialize=uninitialize;uninitialize()
        try:
            deadline=time.monotonic()+120
            for handle,record in zip(owned,records):
                if supervised:record['cleanup']='SUPERVISOR_WAITS_AFTER_WORKER_EXIT'
                else:
                    started=time.monotonic()
                    while win32event.WaitForSingleObject(handle,100)!=win32con.WAIT_OBJECT_0 and time.monotonic()<deadline:pass
                    record['grace_seconds']=time.monotonic()-started;record['exit_code']=win32process.GetExitCodeProcess(handle)
                    record['still_running']=win32event.WaitForSingleObject(handle,0)!=win32con.WAIT_OBJECT_0
                    record['cleanup']='GRACEFUL_EXIT' if not record['still_running'] else 'STOP_OWNED_OFFICE_STILL_RUNNING'
            save(diagnostic,{'schema':'STAGE9_OFFICE_OWNERSHIP_V2','instances':records,'supervised':supervised,
                'deferred_apartment_releases':len(deferred),'owned_orphans':None if supervised else sum(r.get('still_running',False) for r in records)})
            if not supervised and any(r.get('still_running') for r in records):raise RuntimeError('OWNED_OFFICE_DID_NOT_EXIT_NO_FORCE')
        finally:
            for handle in owned:handle.Close()
