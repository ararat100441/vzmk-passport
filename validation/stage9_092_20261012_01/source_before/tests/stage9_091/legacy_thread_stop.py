"""One bounded cleanup attempt for two proven empty legacy TEST COM servers."""
from common import *
import win32api,win32process,win32event,win32gui,ctypes,time
from ctypes import wintypes as W
from stage9_app.ownership import process_image
RUN=OUT/('legacy_thread_stop_'+uuid4().hex[:8]);RUN.mkdir();results=[]
for pid,birth in {43352:'2026-10-11T22:19:00.657000+00:00',68852:'2026-10-11T22:29:18.780000+00:00'}.items():
    matches=[]
    for path in (Path(os.environ['LOCALAPPDATA'])/'VZMKPassport/diagnostics').glob('job_*/office_*_START.json'):
        r=read(path)
        if r.get('pid')==pid and r.get('creation_time')==birth:matches.append((path,r))
    assert len(matches)==1
    path,r=matches[0];h=win32api.OpenProcess(0x100400,False,pid)
    try:
        assert r['owned'] and win32process.GetProcessTimes(h)['CreationTime'].isoformat()==birth and process_image(h)==r['image']
        tid,window_pid=win32process.GetWindowThreadProcessId(r['hwnd']);assert window_pid==pid and win32gui.GetClassName(r['hwnd'])=='XLMAIN'
        thread=win32api.OpenThread(0x100040,False,tid)
        try:
            kernel=ctypes.WinDLL('kernel32',use_last_error=True);getpid=kernel.GetProcessIdOfThread;getpid.argtypes=[W.HANDLE];getpid.restype=W.DWORD
            assert getpid(int(thread))==pid
            proof={'pid':pid,'creation_time':birth,'image':r['image'],'source':str(path),'thread_id':tid,'thread_handle':int(thread),'process_handle':int(h),'action':'WM_QUIT_TO_OWN_LEGACY_EMPTY_TEST_SERVER_THREAD','no_TerminateProcess':True,'monotonic':time.monotonic()}
            save(RUN/f'{pid}_BEFORE.json',proof)
            post=ctypes.WinDLL('user32',use_last_error=True).PostThreadMessageW;post.argtypes=[W.DWORD,W.UINT,W.WPARAM,W.LPARAM];post.restype=W.BOOL
            if not post(tid,0x12,0,0):raise ctypes.WinError(ctypes.get_last_error())
            proof.update(still_running=win32event.WaitForSingleObject(h,30000)==258,exit_code=win32process.GetExitCodeProcess(h));results.append(proof);save(RUN/f'{pid}_AFTER.json',proof)
        finally:thread.Close()
    finally:h.Close()
save(RUN/'RESULT.json',{'instances':results,'still_running':sum(r['still_running'] for r in results)})
print(str(RUN),flush=True)
