from common import *
import time,win32api,win32process,win32gui,win32con,win32event
from stage9_app.ownership import process_image,quit_owned_window
RUN=OUT/('legacy_cleanup_'+uuid4().hex[:8]);RUN.mkdir();results=[]
expected={45184:'2026-10-11T22:09:29.137000+00:00',43352:'2026-10-11T22:19:00.657000+00:00',68852:'2026-10-11T22:29:18.780000+00:00'}
for pid,birth in expected.items():
    candidates=[]
    for path in (Path(os.environ['LOCALAPPDATA'])/'VZMKPassport/diagnostics').glob('job_*/office_*_START.json'):
        r=read(path)
        if r.get('pid')==pid and r.get('creation_time')==birth:candidates.append((path,r))
    assert len(candidates)==1,(pid,candidates)
    source,r=candidates[0];h=win32api.OpenProcess(0x100400,False,pid)
    try:
        assert win32process.GetProcessTimes(h)['CreationTime'].isoformat()==birth and process_image(h)==r['image'] and r['owned'] is True
        record={'pid':pid,'creation_time':birth,'source_proof':str(source),'actions':[]}
        save(RUN/f'{pid}_BEFORE.json',record)
        windows=[];win32gui.EnumWindows(lambda hwnd,_:windows.append(hwnd) if win32process.GetWindowThreadProcessId(hwnd)[1]==pid else None,None)
        if pid==45184:
            for hwnd in windows:
                if win32gui.GetClassName(hwnd)=='#32770':
                    texts=[];win32gui.EnumChildWindows(hwnd,lambda child,_:texts.append(win32gui.GetWindowText(child)),None)
                    if any('Невозможно закрыть приложение Microsoft Word' in s for s in texts):
                        win32gui.SendMessageTimeout(hwnd,win32con.WM_COMMAND,2,0,win32con.SMTO_ABORTIFHUNG,3000);record['actions'].append('dismiss proven own Word cannot-close informational dialog')
            for hwnd in windows:
                if win32gui.IsWindow(hwnd) and win32gui.GetClassName(hwnd)=='NUIDialog':
                    win32gui.SendMessageTimeout(hwnd,win32con.WM_CLOSE,0,0,win32con.SMTO_ABORTIFHUNG,3000);record['actions'].append('cancel own pending save dialog, no save')
            time.sleep(1)
            try:record['quit']=quit_owned_window(r,h)
            except Exception as exc:record['quit_error']=repr(exc)
        else:
            hwnd=r['hwnd']
            assert win32gui.IsWindow(hwnd) and win32process.GetWindowThreadProcessId(hwnd)[1]==pid and win32gui.GetClassName(hwnd)=='XLMAIN'
            win32gui.SendMessageTimeout(hwnd,win32con.WM_CLOSE,0,0,win32con.SMTO_ABORTIFHUNG,3000);record['actions'].append('WM_CLOSE own already-empty Excel root after previous no-save COM Quit')
        record['still_running']=win32event.WaitForSingleObject(h,30000)==258
        record['exit_code']=win32process.GetExitCodeProcess(h);results.append(record);save(RUN/f'{pid}_AFTER.json',record)
    finally:h.Close()
save(RUN/'RESULT.json',{'instances':results,'no_forced_termination':True})
print(str(RUN),flush=True)
