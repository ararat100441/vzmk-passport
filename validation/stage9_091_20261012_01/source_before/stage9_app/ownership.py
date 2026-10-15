"""Track DispatchEx handles for this worker, never attach to existing Office."""
from contextlib import contextmanager
from pathlib import Path
import time
import weakref

class OwnedApplication:
    """Transparent application proxy records the accepted engine's Quit attempt."""
    def __init__(self,app,record,event):object.__setattr__(self,'_app',app);object.__setattr__(self,'_record',record);object.__setattr__(self,'_event',event)
    def __getattr__(self,name):return getattr(self._app,name)
    def __setattr__(self,name,value):setattr(self._app,name,value)
    def Quit(self,*args,**kwargs):
        self._record['quit_attempted']=True
        self._event('QUIT_BEGIN',self._record)
        try:
            result=self._app.Quit(*args,**kwargs);self._record['quit_returned']=True
            self._event('QUIT_RETURNED',self._record);return result
        except BaseException as exc:
            self._record['quit_error']=repr(exc);self._event('QUIT_ERROR',self._record);raise

@contextmanager
def owned_office(diagnostic):
    import win32com.client,win32process,win32api,win32con,win32event,pythoncom,win32gui
    from .workspace import save
    original=win32com.client.DispatchEx
    def event(kind,data):
        from uuid import uuid4
        save(Path(diagnostic).with_name(Path(diagnostic).stem+'_'+kind+'_'+uuid4().hex[:8]+'.json'),{'event':kind,'monotonic':time.monotonic(),'worker_pid':__import__('os').getpid(),'data':dict(data)})
    # Hold a worker apartment across accepted modules' nested Office sessions.
    pythoncom.CoInitialize()
    original_uninitialize=pythoncom.CoUninitialize
    deferred=[]
    # pywin32 CoInitialize is idempotent on its main thread, not reference
    # counted like raw CoInitializeEx. A nested accepted session's uninitialize
    # otherwise disconnects the next module. Keep this worker apartment alive.
    pythoncom.CoUninitialize=lambda:deferred.append('deferred until worker exit')
    before=set(win32process.EnumProcesses());owned=[];records=[]
    def dispatch(progid,*args,**kwargs):
        app=original(progid,*args,**kwargs)
        if str(progid).casefold() in ('excel.application','word.application'):
            if str(progid).casefold()=='excel.application':
                pid=win32process.GetWindowThreadProcessId(app.Hwnd)[1]
            else:
                # Word has no Application.Hwnd before its first document.
                candidates=set()
                def window(hwnd,_):
                    if win32gui.GetClassName(hwnd)=='OpusApp':
                        candidate=win32process.GetWindowThreadProcessId(hwnd)[1]
                        if candidate not in before:candidates.add(candidate)
                for _ in range(20):
                    win32gui.EnumWindows(window,None)
                    if candidates:break
                    time.sleep(.1)
                tracked={r['pid'] for r in records}
                fresh=candidates-tracked
                if len(fresh)!=1:raise RuntimeError('WORD_OWNERSHIP_AMBIGUOUS')
                pid=fresh.pop()
            if pid in before:raise RuntimeError('OFFICE_OWNERSHIP_NOT_PROVEN')
            handle=win32api.OpenProcess(win32con.PROCESS_QUERY_INFORMATION|win32con.PROCESS_TERMINATE|win32con.SYNCHRONIZE,False,pid)
            created=win32process.GetProcessTimes(handle)['CreationTime'].isoformat()
            record={'pid':pid,'creation_time':created,'progid':progid,'owned':True,'quit_attempted':False}
            proxy=OwnedApplication(app,record,event)
            owned.append((weakref.ref(proxy),handle,progid));records.append(record)
            save(Path(diagnostic).with_name(Path(diagnostic).stem+f'_{pid}_START.json'),dict(record))
            return proxy
        return app
    win32com.client.DispatchEx=dispatch
    try:yield records
    finally:
        win32com.client.DispatchEx=original
        # Accepted engine contexts (or Viewer.close) already attempt Close/Quit.
        # A repeated remote call can hang against a server that is shutting down.
        # Do not retain application proxies or repeat these remote calls here.
        app=None
        import gc;gc.collect()
        # Release this worker's apartment before waiting for COM servers.
        # An Office process may already report exit code 0 while its handle
        # cannot signal until the client's apartment releases RPC references.
        pythoncom.CoUninitialize=original_uninitialize
        original_uninitialize()
        # Handles refer to the actual created process and cannot target a reused PID.
        for (_,handle,_),record in zip(owned,records):
            event('BEFORE_WAIT',record|{'exit_code':win32process.GetExitCodeProcess(handle),'com_interfaces':pythoncom._GetInterfaceCount(),'handle':int(handle)})
            if win32event.WaitForSingleObject(handle,3000)!=win32con.WAIT_OBJECT_0 and win32process.GetExitCodeProcess(handle)==259:
                # Worker exclusively owns these invisible instances. Operations
                # have returned or failed; durable Stage 8 recovery handles errors.
                try:win32api.TerminateProcess(handle,0)
                except Exception as exc:
                    # Quit may finish between the timeout and TerminateProcess.
                    # Windows may reject termination after ExitProcess has
                    # begun, before its handle becomes signalled. Give that
                    # already-exiting process time to release COM/RPC threads.
                    record['termination_error']=str(exc)
                    event('TERMINATE_ERROR',record|{'exit_code':win32process.GetExitCodeProcess(handle),'wait':win32event.WaitForSingleObject(handle,0)})
                    if win32event.WaitForSingleObject(handle,10000)!=win32con.WAIT_OBJECT_0 and win32process.GetExitCodeProcess(handle)==259:raise
                if win32event.WaitForSingleObject(handle,10000)!=win32con.WAIT_OBJECT_0 and win32process.GetExitCodeProcess(handle)==259:
                    raise RuntimeError('OWNED_OFFICE_DID_NOT_EXIT')
                record['cleanup']='OWNED_WORKER_PROCESS_TERMINATED_AFTER_QUIT'
            else:record.setdefault('cleanup','QUIT_EXITED')
            record['exit_code']=win32process.GetExitCodeProcess(handle)
            record['still_running']=record['exit_code']==259 and win32event.WaitForSingleObject(handle,0)!=win32con.WAIT_OBJECT_0
            handle.Close()
        owned.clear()
        pythoncom.CoUninitialize=original_uninitialize
        save(diagnostic,{'schema':'STAGE9_OFFICE_OWNERSHIP_V1','instances':records,'deferred_apartment_releases':len(deferred),'owned_orphans':sum(r.get('still_running',False) for r in records)})
