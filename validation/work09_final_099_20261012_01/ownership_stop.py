# Adapted independent run controller; original SHA dc57b2761392d952f8b538fb1a69e0005792258bbe91fab65c79e10916ef62ed
"""A forged Office ownership record must STOP without touching baseline Office."""
from common import *
import ctypes,re,time,win32api,win32process,win32event
from ctypes import wintypes as W
from stage9_app.ownership import process_image
RUN=OUT/('ownership_stop_'+uuid4().hex[:8]);RUN.mkdir()
life=read(read(OUT/'LIFECYCLE_FINAL.json')['evidence'])
users=read(os.environ['STAGE9_FOREIGN_OFFICE']) if os.environ.get('STAGE9_FOREIGN_OFFICE') else life['users']
user_pid=int(next(pid for pid,r in users.items() if Path(r['image']).name.upper()=='EXCEL.EXE'))
user=win32api.OpenProcess(0x100400,False,user_pid);baseline=users[str(user_pid)]
assert win32process.GetProcessTimes(user)['CreationTime'].isoformat()==baseline['creation_time']
result=RUN/'EXE_RESULT.json';parent=subprocess.Popen([str(EXE),'--task','office-probe','--workspace',life['workspace'],'--result',str(result)],env=ENV,creationflags=subprocess.CREATE_NO_WINDOW)
class Entry(ctypes.Structure):
    _fields_=[('size',W.DWORD),('usage',W.DWORD),('pid',W.DWORD),('heap',ctypes.c_size_t),('module',W.DWORD),('threads',W.DWORD),('ppid',W.DWORD),('priority',W.LONG),('flags',W.DWORD),('name',W.WCHAR*260)]
kernel=ctypes.WinDLL('kernel32',use_last_error=True);snapshot=kernel.CreateToolhelp32Snapshot;snapshot.argtypes=[W.DWORD,W.DWORD];snapshot.restype=W.HANDLE
first=kernel.Process32FirstW;nextp=kernel.Process32NextW
for fn in (first,nextp):fn.argtypes=[W.HANDLE,ctypes.POINTER(Entry)];fn.restype=W.BOOL
close=kernel.CloseHandle;close.argtypes=[W.HANDLE];close.restype=W.BOOL
query=ctypes.WinDLL('ntdll').NtQueryInformationProcess;query.argtypes=[W.HANDLE,W.ULONG,ctypes.c_void_p,W.ULONG,ctypes.POINTER(W.ULONG)];query.restype=ctypes.c_long
class Unicode(ctypes.Structure):_fields_=[('length',W.USHORT),('maximum',W.USHORT),('buffer',ctypes.c_void_p)]
deadline=time.monotonic()+10;diagnostic=None
while time.monotonic()<deadline and parent.poll() is None and diagnostic is None:
    snap=snapshot(2,0);entry=Entry();entry.size=ctypes.sizeof(entry)
    try:
        okay=first(snap,ctypes.byref(entry))
        while okay:
            if entry.ppid==parent.pid and entry.name==EXE.name:
                h=win32api.OpenProcess(0x100400,False,entry.pid)
                try:
                    assert process_image(h).casefold()==str(EXE).casefold()
                    needed=W.ULONG();query(int(h),60,None,0,ctypes.byref(needed));buffer=ctypes.create_string_buffer(needed.value)
                    if query(int(h),60,buffer,len(buffer),ctypes.byref(needed))==0:
                        s=Unicode.from_buffer(buffer);command=ctypes.wstring_at(s.buffer,s.length//2)
                        m=re.search(r'--office-worker\s+"?([^"\r\n]+?office\.json)"?(?:\s|$)',command)
                        if m:diagnostic=Path(m.group(1));worker=entry.pid;break
                finally:h.Close()
            okay=nextp(snap,ctypes.byref(entry))
    finally:close(snap)
    if diagnostic is None:time.sleep(.01)
assert diagnostic is not None,'OWN_WORKER_COMMANDLINE_NOT_FOUND'
assert diagnostic.parent.resolve().is_relative_to((Path(SETUP['base'])/'VZMKPassport/diagnostics').resolve())
save(diagnostic.with_name('office_'+str(user_pid)+'_START.json'),{'pid':user_pid,'creation_time':baseline['creation_time'],'image':baseline['image'],'worker_pid':worker,'launcher_pid':parent.pid,'owned':False,'actual_child':False,'absent_before_dispatch':False})
parent.wait(timeout=240);data=read(result);summary=read(data['office_supervisor'])
assert data['status']=='FAIL' and summary['ownership_errors'] and summary['owned_orphans']==0
assert user_pid not in [r['pid'] for r in summary['instances']]
assert win32event.WaitForSingleObject(user,0)==258 and win32process.GetProcessTimes(user)['CreationTime'].isoformat()==baseline['creation_time']
user.Close()
save(RUN/'RESULT.json',{'status':'PASS','exe_sha256':sha(EXE),'worker_pid':worker,'result':str(result),'user_pid':user_pid,'user_process_preserved':True,'unknown_process_not_terminated':True})
save(OUT/'OWNERSHIP_STOP_FINAL.json',{'status':'PASS','exe_sha256':sha(EXE),'evidence':str(RUN/'RESULT.json')})
print('UNKNOWN OWNERSHIP SAFE STOP PASS',flush=True)
