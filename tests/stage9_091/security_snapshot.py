from common import *
import win32api,win32process,win32security,ctypes,json
from ctypes import wintypes as W
from stage9_app.ownership import process_image
records=[]
for pid in [os.getpid(),68852,43352,45184,51208]:
    item={'pid':pid}
    try:
        h=win32api.OpenProcess(0x100400,False,pid)
        try:
            item.update(image=process_image(h),birth=win32process.GetProcessTimes(h)['CreationTime'].isoformat(),exit_code=win32process.GetExitCodeProcess(h))
            token=win32security.OpenProcessToken(h,8)
            try:
                for name,klass in [('user',1),('elevation_type',18),('elevated',20),('integrity',25),('app_container',29)]:
                    try:
                        v=win32security.GetTokenInformation(token,klass)
                        if name=='user':v=win32security.ConvertSidToStringSid(v[0])
                        if name=='integrity':v=win32security.ConvertSidToStringSid(v[0] if isinstance(v,tuple) else v)
                        item[name]=v
                    except Exception as exc:item[name+'_error']=repr(exc)
            finally:token.Close()
            token=win32security.OpenProcessToken(h,8)
            try:
                fn_token=ctypes.WinDLL('advapi32',use_last_error=True).GetTokenInformation
                fn_token.argtypes=[W.HANDLE,ctypes.c_int,ctypes.c_void_p,W.DWORD,ctypes.POINTER(W.DWORD)];fn_token.restype=W.BOOL
                value=W.DWORD();needed=W.DWORD()
                if fn_token(int(token),29,ctypes.byref(value),4,ctypes.byref(needed)):item['app_container_native']=bool(value.value)
                else:item['app_container_native_error']=ctypes.get_last_error()
            finally:token.Close()
            fn=ctypes.WinDLL('ntdll').NtQueryInformationProcess;fn.argtypes=[W.HANDLE,W.ULONG,ctypes.c_void_p,W.ULONG,ctypes.POINTER(W.ULONG)];fn.restype=ctypes.c_long
            protection=ctypes.c_ubyte();length=W.ULONG()
            status=fn(int(h),61,ctypes.byref(protection),1,ctypes.byref(length))
            item.update(protection_status=hex(status&0xffffffff),protection=hex(protection.value))
        finally:h.Close()
    except Exception as exc:item['error']=repr(exc)
    records.append(item)
save(OUT/('SECURITY_SNAPSHOT_'+uuid4().hex[:8]+'.json'),records)
print(json.dumps(records,ensure_ascii=False,default=str,indent=2))
