import ctypes
from pathlib import Path
import win32con
import win32file
from stage8_closure.storage import frozen

r=Path('validation/stage8_20261010_01').resolve()
a=r/'rename_target.txt';b=r/'rename_stage.txt';a.write_text('a');b.write_text('b')
name=str(a)
class Info(ctypes.Structure):
    _fields_=[('Flags',ctypes.c_uint32),('RootDirectory',ctypes.c_void_p),
              ('FileNameLength',ctypes.c_uint32),('FileName',ctypes.c_wchar*len(name))]
info=Info(3,None,len(name.encode('utf-16-le')),name)
kernel=ctypes.WinDLL('kernel32',use_last_error=True)
fn=kernel.SetFileInformationByHandle
fn.argtypes=[ctypes.c_void_p,ctypes.c_int,ctypes.c_void_p,ctypes.c_uint32];fn.restype=ctypes.c_int
with frozen([a],replaceable=a):
    h=win32file.CreateFile(str(b),win32con.GENERIC_READ|win32con.DELETE,
        win32con.FILE_SHARE_READ|win32con.FILE_SHARE_DELETE,None,win32con.OPEN_EXISTING,0,None)
    try:
        if not fn(int(h),22,ctypes.byref(info),ctypes.sizeof(info)):raise ctypes.WinError(ctypes.get_last_error())
    finally:h.Close()
    print('HANDLE_RENAME_PASS',a.read_text())
