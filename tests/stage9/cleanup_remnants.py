"""Clean two hidden Word servers left by documented pre-fix Stage 9 tests.

Exact PID + creation time + hidden empty OpusApp windows are mandatory.
The pre-existing user Excel PID 51208 is outside this operation.
"""
import sys,time
from pathlib import Path
from datetime import datetime
ROOT=Path(__file__).resolve().parents[2];sys.path.insert(0,str(ROOT))
from stage9_app.workspace import save
import win32api,win32con,win32process,win32event,win32gui
expected={31492:'2026-10-12T00:57:28.3867926+04:00',65120:'2026-10-12T01:09:55.4375891+04:00'}
results=[]
for pid,birth in expected.items():
    try:h=win32api.OpenProcess(win32con.PROCESS_QUERY_INFORMATION|win32con.PROCESS_TERMINATE|win32con.SYNCHRONIZE,False,pid)
    except Exception:results.append({'pid':pid,'status':'ALREADY_EXITED'});continue
    try:
        actual=win32process.GetProcessTimes(h)['CreationTime']
        assert abs(actual.timestamp()-datetime.fromisoformat(birth).timestamp())<.001,(pid,actual)
        windows=[]
        def enum(hwnd,_):
            if win32process.GetWindowThreadProcessId(hwnd)[1]==pid and win32gui.GetClassName(hwnd)=='OpusApp':windows.append(hwnd)
        win32gui.EnumWindows(enum,None)
        assert windows and all(not win32gui.IsWindowVisible(w) and win32gui.GetWindowText(w) in ('','Word') for w in windows)
        for w in windows:win32gui.PostMessage(w,win32con.WM_CLOSE,0,0)
        if win32event.WaitForSingleObject(h,3000)!=win32con.WAIT_OBJECT_0:win32api.TerminateProcess(h,0)
        assert win32event.WaitForSingleObject(h,3000)==win32con.WAIT_OBJECT_0
        results.append({'pid':pid,'creation_time':actual.isoformat(),'status':'OWNED_PRE_FIX_TEST_SERVER_CLOSED','hidden_empty_windows':True})
    finally:h.Close()
save(ROOT/'validation/stage9_20261012_01/PRE_FIX_OFFICE_CLEANUP.json',results)
