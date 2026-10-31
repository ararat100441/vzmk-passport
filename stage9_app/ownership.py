"""Bind DispatchEx to actual Office windows; graceful teardown, no global kill."""
from contextlib import contextmanager
from pathlib import Path
import os
import time
import sys
from threading import Lock

_guard_lock=Lock()

def wait_owned_exit(handles,timeout):
    """Bounded graceful wait with the client's STA alive; never terminate Office."""
    import pythoncom,win32event,win32con
    deadline=time.monotonic()+timeout
    while any(win32event.WaitForSingleObject(h,0)!=win32con.WAIT_OBJECT_0 for h in handles):
        if time.monotonic()>=deadline:return False
        pythoncom.PumpWaitingMessages();time.sleep(.05)
    return True

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
    # Excel's registered COM server uses automation + COM's -Embedding flag.
    # Do not mix the interactive /x startup mode with automation.
    # One documented Word switch: a private startup ignores global add-ins,
    # Normal.dotm and user settings. The two-switch /w /m startup could leave
    # an interactive startup dialog before any _WwG native object existed.
    # PID/creation-time/retained-handle checks still prove exclusive ownership.
    switches=['/a'] if word else ['/automation','-Embedding']
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
            'retained_handle':int(child._handle),'handle_source':'CREATE_PROCESS',
            'startup_switches':switches}
    if record['image'].casefold()!=image.casefold():raise RuntimeError('OFFICE_CHILD_IMAGE_CHANGED')
    return child,record

def startup_problem(windows,expected):
    """Explain a real modal; do not misdiagnose a slow splash as a dialog."""
    dialogs=[w for w in windows if w.get('class')=='#32770' and w.get('visible')]
    if not dialogs:return ''
    texts=[]
    for window in dialogs:
        texts.append(window.get('title',''))
        texts.extend(c.get('text','') for c in window.get('children',[]) if c.get('text'))
    return 'Office: рабочее окно '+expected+' не создано; запуск остановил диалог: '+' | '.join(dict.fromkeys(texts))

def owned_window_snapshot(record):
    """Local diagnostics, restricted to the retained process's PID."""
    import win32gui,win32process
    windows=[]
    def visit(hwnd,_):
        if win32process.GetWindowThreadProcessId(hwnd)[1]!=record['pid']:return
        info={'hwnd':hwnd,'class':win32gui.GetClassName(hwnd),
              'title':win32gui.GetWindowText(hwnd),'visible':bool(win32gui.IsWindowVisible(hwnd)),'children':[]}
        def child(ch,_):
            if win32process.GetWindowThreadProcessId(ch)[1]==record['pid']:
                info['children'].append({'hwnd':ch,'class':win32gui.GetClassName(ch),'text':win32gui.GetWindowText(ch)[:500]})
        win32gui.EnumChildWindows(hwnd,child,None);windows.append(info)
    win32gui.EnumWindows(visit,None)
    return windows

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
    if not candidates:
        import win32event,win32process
        snapshot=owned_window_snapshot(record)
        trace('NATIVEOM_NOT_READY',{'pid':record['pid'],'creation_time':record['creation_time'],
                                  'expected_class':expected,'windows':snapshot})
        if win32event.WaitForSingleObject(handle,0)==0:
            raise RuntimeError('Собственный процесс Office завершился до подключения; код '+str(win32process.GetExitCodeProcess(handle)))
        raise RuntimeError(startup_problem(snapshot,expected) or 'OWNED_NATIVEOM_WINDOW_MISSING')
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
                record['workbooks_before_quit']=int(app.Workbooks.Count)
                app.Quit()
            record['quit_returned']=True
            trace('QUIT_RETURNED',record['pid'])
        app=None
        import win32event,win32con,win32process
        deadline=time.monotonic()+15
        while win32event.WaitForSingleObject(handle,100)!=win32con.WAIT_OBJECT_0 and time.monotonic()<deadline:pythoncom.PumpWaitingMessages()
        if win32event.WaitForSingleObject(handle,0)!=win32con.WAIT_OBJECT_0:
            request_empty_excel_close(record,handle,trace)
            deadline=time.monotonic()+15
            while win32event.WaitForSingleObject(handle,100)!=win32con.WAIT_OBJECT_0 and time.monotonic()<deadline:pythoncom.PumpWaitingMessages()
        exited=win32event.WaitForSingleObject(handle,0)==win32con.WAIT_OBJECT_0
        return {'status':'PASS' if exited else 'FAIL','action':'NATIVEOM_CLOSE_WITHOUT_SAVE_AND_QUIT','pid':record['pid'],'creation_time':record['creation_time'],'exit_observed_while_sta_alive':exited,'exit_code':win32process.GetExitCodeProcess(handle)}
    finally:pythoncom.CoUninitialize()

def request_empty_excel_close(record,handle,trace):
    """Normal cancellable window close, only after an owned empty Excel Quit.

    WM_CLOSE follows Excel's own shutdown procedure. Never post WM_QUIT, and
    never close a workbook or a window belonging to an unproven process here.
    """
    import pythoncom,win32process,win32gui
    if Path(record['image']).name.upper()!='EXCEL.EXE':return
    if not record.get('quit_returned') or record.get('workbooks_before_quit')!=0:return
    if pythoncom._GetInterfaceCount()!=0:raise RuntimeError('COM_REFERENCES_REMAIN_NO_WINDOW_CLOSE')
    if (record.get('actual_child') is not True or record.get('owned') is not True
        or record.get('handle_source')!='CREATE_PROCESS'
        or record.get('launcher_pid') not in (os.getpid(),os.getppid())
        or win32process.GetProcessId(handle)!=record['pid']
        or win32process.GetProcessTimes(handle)['CreationTime'].isoformat()!=record['creation_time']
        or process_image(handle)!=record['image']):raise RuntimeError('OFFICE_IDENTITY_CHANGED_NO_WINDOW_CLOSE')
    windows=[]
    def window(hwnd,_):
        if win32process.GetWindowThreadProcessId(hwnd)[1]==record['pid'] and win32gui.GetClassName(hwnd)=='XLMAIN':windows.append(hwnd)
    win32gui.EnumWindows(window,None)
    trace('EMPTY_EXCEL_WM_CLOSE_BEGIN',dict(record,windows=windows,handle_validated=True,com_interfaces=0,message='WM_CLOSE'))
    for hwnd in windows:
        if win32gui.IsWindow(hwnd) and win32process.GetWindowThreadProcessId(hwnd)[1]==record['pid'] and win32gui.GetClassName(hwnd)=='XLMAIN':
            win32gui.PostMessage(hwnd,0x0010,0,0)
    record['normal_window_close_requested']=bool(windows)
    trace('EMPTY_EXCEL_WM_CLOSE_POSTED',record)

class OwnedApplication:
    def __init__(self,app,record,event):
        object.__setattr__(self,'_app',app);object.__setattr__(self,'_record',record);object.__setattr__(self,'_event',event)
    def __getattr__(self,name):return getattr(self._app,name)
    def __setattr__(self,name,value):setattr(self._app,name,value)
    def Quit(self,*args,**kwargs):
        if str(self._record['progid']).casefold()=='excel.application':
            try:
                self._record['workbooks_before_quit']=int(self._app.Workbooks.Count)
                self._record['user_control_before_quit']=bool(self._app.UserControl)
            except Exception as exc:self._record['quit_state_diagnostic_error']=repr(exc)
        self._record['quit_attempted']=True;self._event('QUIT_BEGIN',self._record)
        try:
            value=self._app.Quit(*args,**kwargs)
            self._record['quit_returned']=True;self._event('QUIT_RETURNED',self._record)
            return value
        except BaseException as exc:
            self._record['quit_error']=repr(exc);self._event('QUIT_ERROR',self._record);raise

@contextmanager
def _owned_office(diagnostic, *, supervised=False, cleanup_errors=None):
    import win32com.client,win32process,win32api,win32con,win32event,pythoncom
    from .workspace import save
    from .diagnostics import event as save_event,error_record,OfficeCleanupError
    from uuid import uuid4
    diagnostic=Path(diagnostic)
    diagnostic.parent.mkdir(parents=True,exist_ok=True)
    errors=cleanup_errors if cleanup_errors is not None else []
    diagnostic_errors=[]
    def event(kind,data):
        try:
            save_event(diagnostic,kind,{'event':kind,'monotonic':time.monotonic(),
                       'worker_pid':os.getpid(),'data':dict(data)})
            return True
        except Exception as exc:
            diagnostic_errors.append(error_record('Diagnostic.'+kind,exc))
            return False
    original=win32com.client.DispatchEx;initialize=pythoncom.CoInitialize;initialize()
    uninitialize=pythoncom.CoUninitialize;deferred=[]
    nested_initializations=[]
    # One guard owns the STA. Accepted contexts may enter/leave it repeatedly;
    # defer both sides, rather than accumulating unmatched CoInitialize calls.
    pythoncom.CoInitialize=lambda:nested_initializations.append(True)
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
                with native_application(record,handle,lambda phase,details:event(phase,{'pid':record['pid'],'details':details})) as native:app=native
                native=None;break
            except RuntimeError as exc:
                if str(exc)!='OWNED_NATIVEOM_WINDOW_MISSING' or time.monotonic()>=deadline:raise
                time.sleep(.1)
        app.AutomationSecurity=3;app.DisplayAlerts=0 if suffix=='.docx' else False;app.Visible=False
        if suffix=='.xlsx':
            # CreateProcess starts a user-controlled Excel, unlike CreateObject.
            # This exclusively owned hidden server must also exit when its last
            # COM reference is released, including after a worker crash.
            record['user_control_before_binding']=bool(app.UserControl)
            app.UserControl=False
            record['user_control_after_binding']=bool(app.UserControl)
            if record['user_control_after_binding']:raise RuntimeError('OWNED_EXCEL_AUTOMATION_LIFETIME_NOT_SET')
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
        primary=sys.exc_info()[1]
        win32com.client.DispatchEx=original
        import gc
        try:
            interfaces_before_gc=pythoncom._GetInterfaceCount();gc.collect()
            event('COM_RELEASE_STATE',{'interfaces_before_gc':interfaces_before_gc,
                  'interfaces_after_gc':pythoncom._GetInterfaceCount(),'deferred_apartment_releases':len(deferred),
                  'nested_apartment_initializations':len(nested_initializations),'guard_initializations':1,'guard_releases':1})
            # Excel Quit can require the client's STA to service final calls.
            # Keep that apartment alive after releasing proxies until process
            # exit is observed; a dead worker is too late for this handshake.
            wait_owned_exit(owned,15)
            for handle,record in zip(owned,records):
                if win32event.WaitForSingleObject(handle,0)!=win32con.WAIT_OBJECT_0:request_empty_excel_close(record,handle,event)
            # Office's background shutdown (including its power-manager window)
            # can take almost a minute after Quit with zero COM interfaces and
            # zero workbooks. Observe that exit in the live STA instead of
            # prematurely ending the worker; ownership/exit-code checks below
            # remain mandatory and the wait is still bounded.
            wait_owned_exit(owned,60)
            for handle,record in zip(owned,records):
                record['exit_code']=win32process.GetExitCodeProcess(handle)
                record['still_running']=win32event.WaitForSingleObject(handle,0)!=win32con.WAIT_OBJECT_0
                record['exit_observed_while_sta_alive']=not record['still_running']
                record['cleanup']='GRACEFUL_EXIT_WITH_LIVE_STA' if not record['still_running'] and record['exit_code']==0 else 'STOP_OWNED_OFFICE_STILL_RUNNING_OR_ABNORMAL'
            if any(r.get('still_running') or r['exit_code']!=0 for r in records):raise RuntimeError('OWNED_OFFICE_STA_GRACE_TIMEOUT_OR_ABNORMAL_EXIT')
        except Exception as exc:errors.append(error_record('Office.GracefulTeardown',exc))
        finally:
            pythoncom.CoInitialize=initialize;pythoncom.CoUninitialize=uninitialize
            try:uninitialize()
            except Exception as exc:errors.append(error_record('COM.CoUninitialize',exc))
            try:
                if not supervised:
                    for handle,record in zip(owned,records):
                        try:
                            if win32event.WaitForSingleObject(handle,0)==win32con.WAIT_OBJECT_0:continue
                            if (record.get('actual_child') is not True or record.get('owned') is not True
                                or record.get('launcher_pid')!=os.getpid()
                                or record.get('handle_source')!='CREATE_PROCESS' or int(handle)!=record['retained_handle']
                                or win32process.GetProcessId(handle)!=record['pid']
                                or win32process.GetProcessTimes(handle)['CreationTime'].isoformat()!=record['creation_time']
                                or process_image(handle)!=record['image']):raise RuntimeError('VIEWER_IDENTITY_CHANGED_NO_TERMINATE')
                            if not event('VIEWER_EMERGENCY_FORCE_BEGIN',dict(record,handle_validated=True)):
                                raise RuntimeError('FORCE_DIAGNOSTIC_UNAVAILABLE_NO_TERMINATE')
                            record['forced_termination_attempted']=True
                            win32api.TerminateProcess(handle,1);win32event.WaitForSingleObject(handle,10000)
                            record['exit_code']=win32process.GetExitCodeProcess(handle)
                            record['still_running']=win32event.WaitForSingleObject(handle,0)!=win32con.WAIT_OBJECT_0
                            record['cleanup']='ABNORMAL_OR_FORCED_EXIT'
                            event('VIEWER_EMERGENCY_FORCE_RESULT',record)
                        except Exception as exc:errors.append(error_record('Viewer.EmergencyTeardown',exc))
            finally:
                # Popen owns the same CreateProcess handle in direct/native
                # calls. Cache exit before closing it, as the supervisor does;
                # otherwise Popen.__del__ polls an already closed handle.
                for child in processes:
                    try:child.poll()
                    except Exception as exc:errors.append(error_record('Process.ExitCache',exc))
                live_handles={int(child._handle) for child in processes if child.returncode is None}
                for handle in owned:
                    if int(handle) in live_handles:continue  # Popen retains its live handle
                    try:handle.Close()
                    except Exception as exc:errors.append(error_record('ProcessHandle.Close',exc))
        payload={'schema':'STAGE9_OFFICE_OWNERSHIP_V2','instances':records,'supervised':supervised,
                 'deferred_apartment_releases':len(deferred),
                 'owned_orphans':sum(r.get('still_running',True) for r in records),
                 'cleanup_errors':errors,'diagnostic_errors':diagnostic_errors,
                 'primary_error':error_record('Primary',primary) if primary is not None else None,
                 'status':'FAIL' if primary is not None or errors or diagnostic_errors else 'PASS'}
        # One final exclusive write, after emergency outcomes and handle release.
        try:save(diagnostic,payload)
        except Exception as exc:diagnostic_errors.append(error_record('Diagnostic.FinalSave',exc))
        failures=errors+diagnostic_errors
        if failures:
            if primary is not None:
                primary.stage9_cleanup_errors=list(failures)
                primary.add_note(str(OfficeCleanupError(diagnostic,failures)))
            else:raise OfficeCleanupError(diagnostic,failures)

@contextmanager
def owned_office(diagnostic, *, supervised=False, cleanup_errors=None):
    # DispatchEx/apartment hooks are process globals: reject overlapping guards.
    # Separate EXE workers remain independently usable in parallel.
    if not _guard_lock.acquire(blocking=False):raise RuntimeError('OFFICE_GUARD_ALREADY_ACTIVE')
    try:
        with _owned_office(diagnostic,supervised=supervised,cleanup_errors=cleanup_errors) as records:yield records
    finally:_guard_lock.release()
