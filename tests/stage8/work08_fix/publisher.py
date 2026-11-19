"""Independent external process: does not honor the Stage8 application lock."""
import json,os,sys,ctypes
from pathlib import Path
import win32file,win32con

mode,source,target=sys.argv[1:4]
try:
    if mode=='rename':os.rename(target,target+'.external_old');os.replace(source,target)
    elif mode=='replace':os.replace(source,target)
    elif mode=='write':
        with open(target,'r+b') as f:f.seek(0);f.write(b'EXTERNAL_DATA');f.flush();os.fsync(f.fileno())
    else:
        name=str(Path(target).resolve())
        class Info(ctypes.Structure):
            _fields_=[('Flags',ctypes.c_uint32),('RootDirectory',ctypes.c_void_p),
                      ('FileNameLength',ctypes.c_uint32),('FileName',ctypes.c_wchar*len(name))]
        info=Info(3,None,len(name.encode('utf-16-le')),name)
        fn=ctypes.WinDLL('kernel32',use_last_error=True).SetFileInformationByHandle
        fn.argtypes=[ctypes.c_void_p,ctypes.c_int,ctypes.c_void_p,ctypes.c_uint32];fn.restype=ctypes.c_int
        h=win32file.CreateFile(source,win32con.GENERIC_READ|win32con.DELETE,
            win32con.FILE_SHARE_READ|win32con.FILE_SHARE_DELETE,None,win32con.OPEN_EXISTING,0,None)
        try:
            if not fn(int(h),22,ctypes.byref(info),ctypes.sizeof(info)):raise ctypes.WinError(ctypes.get_last_error())
        finally:h.Close()
    print(json.dumps({'status':'PUBLISHED','mode':mode}));sys.exit(0)
except Exception as exc:
    print(json.dumps({'status':'BLOCKED','mode':mode,'error':str(exc),'winerror':getattr(exc,'winerror',None)}));sys.exit(2)
