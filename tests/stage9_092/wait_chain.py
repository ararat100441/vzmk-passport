"""Passive WCT sampling of a validated, still-live owned Office window only."""
from common import *
import ctypes,time,win32api,win32process,win32gui,win32event
from ctypes import wintypes as W
from stage9_app.ownership import process_image
class Lock(ctypes.Structure):_fields_=[('name',W.WCHAR*128),('timeout',ctypes.c_longlong),('alertable',W.BOOL)]
class Thread(ctypes.Structure):_fields_=[('pid',W.DWORD),('tid',W.DWORD),('wait_ms',W.DWORD),('switches',W.DWORD)]
class Details(ctypes.Union):_fields_=[('lock',Lock),('thread',Thread)]
class Node(ctypes.Structure):_fields_=[('kind',ctypes.c_int),('status',ctypes.c_int),('details',Details)]
api=ctypes.WinDLL('advapi32',use_last_error=True)
api.OpenThreadWaitChainSession.argtypes=[W.DWORD,ctypes.c_void_p];api.OpenThreadWaitChainSession.restype=W.HANDLE
api.CloseThreadWaitChainSession.argtypes=[W.HANDLE]
api.GetThreadWaitChain.argtypes=[W.HANDLE,ctypes.c_size_t,W.DWORD,W.DWORD,ctypes.POINTER(W.DWORD),ctypes.POINTER(Node),ctypes.POINTER(W.BOOL)];api.GetThreadWaitChain.restype=W.BOOL
session=api.OpenThreadWaitChainSession(0,None);results=[]
try:
    for job in (Path(SETUP['base'])/'VZMKPassport/diagnostics').glob('job_*'):
        if (job/'SUPERVISOR.json').exists():continue
        for path in job.glob('office_*_START.json'):
            r=read(path);h=None
            try:
                h=win32api.OpenProcess(0x100400,False,r['pid'])
                if win32process.GetProcessTimes(h)['CreationTime'].isoformat()!=r['creation_time'] or process_image(h)!=r['image'] or win32event.WaitForSingleObject(h,0)!=258:continue
                tid,pid=win32process.GetWindowThreadProcessId(r['hwnd'])
                if pid!=r['pid']:continue
                count=W.DWORD(16);nodes=(Node*16)();cycle=W.BOOL()
                okay=api.GetThreadWaitChain(session,0,0,tid,ctypes.byref(count),nodes,ctypes.byref(cycle));chain=[]
                if okay:
                    for node in nodes[:count.value]:
                        item={'kind':node.kind,'status':node.status}
                        if node.kind==8:item.update(pid=node.details.thread.pid,tid=node.details.thread.tid,wait_ms=node.details.thread.wait_ms)
                        else:item['object']=node.details.lock.name
                        chain.append(item)
                results.append({'record':str(path),'pid':pid,'tid':tid,'identity_validated':True,'success':bool(okay),'error':None if okay else ctypes.get_last_error(),'cycle':bool(cycle.value),'chain':chain})
            except Exception as exc:results.append({'record':str(path),'error':repr(exc)})
            finally:
                if h is not None:h.Close()
finally:api.CloseThreadWaitChainSession(session)
save(OUT/('WAIT_CHAIN_'+uuid4().hex[:8]+'.json'),{'passive_only':True,'samples':results})
print('Owned window WCT samples',len(results))
