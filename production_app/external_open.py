"""Launch the user's viewer without leaking PyInstaller's private DLL path.

Matches the native Office launcher isolation. Restore this process before
the next internal worker; never alter Windows or viewer settings.
"""
from contextlib import contextmanager
from pathlib import Path
import os
import sys


@contextmanager
def system_loader():
    if sys.platform!='win32' or not getattr(sys,'frozen',False):
        yield
        return
    import ctypes
    from ctypes import wintypes as W
    kernel=ctypes.WinDLL('kernel32',use_last_error=True)
    get=kernel.GetDllDirectoryW;get.argtypes=[W.DWORD,W.LPWSTR];get.restype=W.DWORD
    setdir=kernel.SetDllDirectoryW;setdir.argtypes=[W.LPCWSTR];setdir.restype=W.BOOL
    buffer=ctypes.create_unicode_buffer(32768)
    size=get(len(buffer),buffer)
    if size>=len(buffer):raise OSError('Не удалось сохранить путь библиотек приложения.')
    previous=buffer.value or None
    original_path=os.environ.get('PATH')
    bundle=os.path.normcase(os.path.abspath(sys._MEIPASS))
    clean=[]
    for part in (original_path or '').split(os.pathsep):
        location=os.path.normcase(os.path.abspath(part.strip('"')))
        if location!=bundle and not location.startswith(bundle+os.sep):clean.append(part)
    if not setdir(None):raise ctypes.WinError(ctypes.get_last_error())
    try:
        os.environ['PATH']=os.pathsep.join(clean)
        yield
    finally:
        if original_path is None:os.environ.pop('PATH',None)
        else:os.environ['PATH']=original_path
        if not setdir(previous):raise ctypes.WinError(ctypes.get_last_error())


def open_external(path):
    with system_loader():os.startfile(path)
