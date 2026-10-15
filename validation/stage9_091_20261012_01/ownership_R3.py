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

def quit_owned_window(record,handle,trace=lambda *args:None):
    """Recover only this proven Office instance through OBJID_NATIVEOM."""
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
    pythoncom.CoInitialize();obj=native=app=None
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
        trace('OWN_APPLICATION_BOUND',record['pid'])
        if expected=='_WwG':app.Quit(SaveChanges=0)
        else:
            app.DisplayAlerts=False
            while app.Workbooks.Count:app.Workbooks(1).Close(SaveChanges=False)
            app.Quit()
        trace('QUIT_RETURNED',record['pid'])
        return {'status':'PASS','action':'NATIVEOM_CLOSE_WITHOUT_SAVE_AND_QUIT','pid':record['pid'],'creation_time':record['creation_time']}
    finally:app=native=obj=None;gc.collect();pythoncom.CoUninitialize()

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
    owned=[];records=[]
    def dispatch(progid,*args,**kwargs):
        before=set(win32process.EnumProcesses());app=original(progid,*args,**kwargs)
        kind=str(progid).casefold()
        if kind not in ('excel.application','word.application'):return app
        probe=None
        if kind=='word.application':
            # A window created through this COM object proves its PID.
            if app.Documents.Count:raise RuntimeError('WORD_OWNERSHIP_NOT_PROVEN')
            app.AutomationSecurity=3;app.DisplayAlerts=0
            probe=app.Documents.Add(Visible=False);hwnd=probe.Windows(1).Hwnd
        else:hwnd=app.Hwnd
        pid=win32process.GetWindowThreadProcessId(hwnd)[1]
        if pid in before:raise RuntimeError('OFFICE_OWNERSHIP_NOT_PROVEN_NO_QUIT')
        handle=win32api.OpenProcess(win32con.PROCESS_QUERY_INFORMATION|win32con.SYNCHRONIZE,False,pid)
        try:
            image=process_image(handle);expected='EXCEL.EXE' if kind=='excel.application' else 'WINWORD.EXE'
            if Path(image).name.upper()!=expected:raise RuntimeError('OFFICE_IMAGE_MISMATCH')
            record={'pid':pid,'creation_time':win32process.GetProcessTimes(handle)['CreationTime'].isoformat(),
                'progid':progid,'image':image,'hwnd':int(hwnd),'worker_pid':os.getpid(),
                'absent_before_dispatch':True,'owned':True,'quit_attempted':False}
            records.append(record);owned.append(handle)
            save(diagnostic.with_name(diagnostic.stem+f'_{pid}_START.json'),dict(record))
        except BaseException:handle.Close();raise
        if probe is not None:probe.Close(False);probe=None
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
