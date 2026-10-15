"""Read-only Office observer: QUERY/SYNCHRONIZE handles; no termination API."""
import os,sys,json,time,ctypes
from pathlib import Path
import win32api,win32process,win32event
OUT=Path(__file__).resolve().parent;ROOT=OUT.parents[1];sys.path.insert(0,str(ROOT))
from stage9_app.ownership import process_image
SETUP=json.loads((OUT/'SETUP.json').read_text(encoding='utf8'));BASE=Path(SETUP['base'])/'VZMKPassport/diagnostics'
def read(p):return json.loads(Path(p).read_text(encoding='utf8'))
def save(p,data):
    with Path(p).open('x',encoding='utf8') as f:json.dump(data,f,ensure_ascii=False,indent=2,default=str);f.flush();os.fsync(f.fileno())
from ctypes import wintypes as W
class Entry(ctypes.Structure):
    _fields_=[('size',W.DWORD),('usage',W.DWORD),('pid',W.DWORD),('heap',ctypes.c_size_t),('module',W.DWORD),('threads',W.DWORD),('ppid',W.DWORD),('priority',W.LONG),('flags',W.DWORD),('name',W.WCHAR*260)]
k=ctypes.WinDLL('kernel32',use_last_error=True)
k.CreateToolhelp32Snapshot.argtypes=[W.DWORD,W.DWORD];k.CreateToolhelp32Snapshot.restype=W.HANDLE
for fn in (k.Process32FirstW,k.Process32NextW):fn.argtypes=[W.HANDLE,ctypes.POINTER(Entry)];fn.restype=W.BOOL
k.CloseHandle.argtypes=[W.HANDLE];k.CloseHandle.restype=W.BOOL
def office_pids():
    result={};snap=k.CreateToolhelp32Snapshot(2,0);e=Entry();e.size=ctypes.sizeof(e)
    try:
        okay=k.Process32FirstW(snap,ctypes.byref(e))
        while okay:
            if e.name.upper() in ('EXCEL.EXE','WINWORD.EXE'):result[e.pid]=e.ppid
            okay=k.Process32NextW(snap,ctypes.byref(e))
    finally:k.CloseHandle(snap)
    return result
def own_parent(pid):
    for p in OUT.glob('CONTROLLER_*_IDENTITY.json'):
        try:
            r=read(p)
            if r['pid']==pid:return True
        except (OSError,ValueError):pass
    for p in OUT.glob('GUI*LAUNCH*.json'):
        try:
            if read(p)['pid']==pid:return True
        except (OSError,ValueError,KeyError):pass
    for p in BASE.glob('job_*/JOB_REQUEST.json'):
        try:
            r=read(p)
            if r['supervisor_pid']==pid:return True
        except (OSError,ValueError,KeyError):pass
    return False
records={};handles={};first=True;stops=[]
log=(OUT/'OFFICE_MONITOR.jsonl').open('x',encoding='utf8')
try:
    while True:
        for pid,parent in office_pids().items():
            if pid in handles:continue
            try:
                h=win32api.OpenProcess(0x1000|0x100000,False,pid)
                if win32event.WaitForSingleObject(h,0)==0:h.Close();continue
                image=process_image(h);birth=win32process.GetProcessTimes(h)['CreationTime'].isoformat()
            except Exception:
                try:h.Close()
                except Exception:pass
                continue
            if Path(image).name.upper() not in ('WINWORD.EXE','EXCEL.EXE'):h.Close();continue
            r={'pid':pid,'parent_pid':parent,'image':image,'creation_time':birth,'baseline':first,
               'own_parent_observed':False if first else own_parent(parent),'first_seen':time.time(),'exit_observed':False}
            records[pid]=r;handles[pid]=h;log.write(json.dumps({'event':'SEEN','identity':r},ensure_ascii=False)+'\n');log.flush()
        if first:
            save(OUT/'OFFICE_BASELINE.json',{'office':list(records.values()),'handle_access':'QUERY_LIMITED_INFORMATION|SYNCHRONIZE','termination_API':False})
            first=False;print('Office monitor ready; baseline',len(records),flush=True)
        for pid,h in handles.items():
            r=records[pid]
            if not r['baseline'] and not r['own_parent_observed']:r['own_parent_observed']=own_parent(r['parent_pid'])
            if not r['exit_observed'] and win32event.WaitForSingleObject(h,0)==0:
                r.update(exit_observed=True,exit_code=win32process.GetExitCodeProcess(h),exit_seen=time.time())
                log.write(json.dumps({'event':'EXIT','identity':r},ensure_ascii=False)+'\n');log.flush()
                if r['baseline'] or not r['own_parent_observed']:
                    stops.append(dict(r))
                    if not (OUT/'FOREIGN_OFFICE_STOP.json').exists():save(OUT/'FOREIGN_OFFICE_STOP.json',{'status':'STOP','unknown_office_exit':r,'cause':'UNKNOWN','no_termination_API':True})
        if (OUT/'STOP_MONITOR.json').exists():break
        time.sleep(.25)
    save(OUT/'OFFICE_MONITOR_FINAL.json',{'instances':list(records.values()),'foreign_stop':stops,'termination_API':False})
finally:
    log.close()
    for h in handles.values():h.Close()
