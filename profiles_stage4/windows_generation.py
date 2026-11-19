"""Windows generation commit: deny external write/delete; rename through owned handles.

Renames NEVER replace a destination. If an external file appears in a temporarily
vacant name, publication stops, preserves that file, and rolls back the owned
files. A hash check followed by ordinary os.replace would not provide this safety.
"""
from contextlib import ExitStack
import ctypes
from ctypes import wintypes
import hashlib
import os
from pathlib import Path

from .model import ProfileError


def digest(data):
    return hashlib.sha256(data).hexdigest()


class ProtectedFile:
    """OPEN_EXISTING + READ/DELETE access, FILE_SHARE_READ only, until closed."""
    def __init__(self, path):
        self.path = Path(path)
        self.handle = None

    def __enter__(self):
        if os.name != 'nt':
            raise ProfileError('Generation commit requires Windows file-handle protection')
        self.api = ctypes.WinDLL('kernel32', use_last_error=True)
        self.api.CreateFileW.argtypes = [wintypes.LPCWSTR, wintypes.DWORD, wintypes.DWORD,
            wintypes.LPVOID, wintypes.DWORD, wintypes.DWORD, wintypes.HANDLE]
        self.api.CreateFileW.restype = wintypes.HANDLE
        self.api.CloseHandle.argtypes = [wintypes.HANDLE]
        self.api.CloseHandle.restype = wintypes.BOOL
        self.api.SetFilePointerEx.argtypes = [wintypes.HANDLE, ctypes.c_longlong, wintypes.LPVOID, wintypes.DWORD]
        self.api.SetFilePointerEx.restype = wintypes.BOOL
        self.api.ReadFile.argtypes = [wintypes.HANDLE, wintypes.LPVOID, wintypes.DWORD,
                                     ctypes.POINTER(wintypes.DWORD), wintypes.LPVOID]
        self.api.ReadFile.restype = wintypes.BOOL
        self.api.SetFileInformationByHandle.argtypes = [wintypes.HANDLE, ctypes.c_int, wintypes.LPVOID, wintypes.DWORD]
        self.api.SetFileInformationByHandle.restype = wintypes.BOOL
        # DELETE access permits our handle-based rename. No FILE_SHARE_WRITE/DELETE
        # permits another process to modify, remove, or replace this file.
        handle = self.api.CreateFileW(str(self.path), 0x80000000 | 0x00010000, 1, None, 3, 0x80, None)
        if handle == ctypes.c_void_p(-1).value:
            raise ProfileError(f'Cannot protect generation file {self.path.name}: {ctypes.WinError(ctypes.get_last_error())}')
        self.handle = handle
        return self

    def read(self):
        if not self.api.SetFilePointerEx(self.handle, 0, None, 0):
            raise ctypes.WinError(ctypes.get_last_error())
        chunks, buffer, count = [], ctypes.create_string_buffer(1024 * 1024), wintypes.DWORD()
        while True:
            if not self.api.ReadFile(self.handle, buffer, len(buffer), ctypes.byref(count), None):
                raise ctypes.WinError(ctypes.get_last_error())
            if not count.value:
                return b''.join(chunks)
            chunks.append(buffer.raw[:count.value])

    def rename(self, destination):
        """Rename by our DELETE-capable handle, with ReplaceIfExists=False."""
        class RenameInfo(ctypes.Structure):
            _fields_ = [('ReplaceIfExists', wintypes.BOOL), ('RootDirectory', wintypes.HANDLE),
                        ('FileNameLength', wintypes.DWORD), ('FileName', wintypes.WCHAR * 1)]
        destination = Path(destination).resolve()
        encoded = str(destination).encode('utf-16-le')
        # Win32's absolute-path conversion also requires a terminating WCHAR,
        # although FileNameLength itself excludes it. Include structure padding.
        size = ctypes.sizeof(RenameInfo) + len(encoded) + ctypes.sizeof(wintypes.WCHAR)
        buffer = ctypes.create_string_buffer(size)
        info = RenameInfo.from_buffer(buffer)
        info.ReplaceIfExists, info.RootDirectory, info.FileNameLength = False, None, len(encoded)
        ctypes.memmove(ctypes.addressof(buffer) + RenameInfo.FileName.offset, encoded, len(encoded))
        if not self.api.SetFileInformationByHandle(self.handle, 3, buffer, size):
            raise ProfileError(f'Protected rename refused {destination.name}: {ctypes.WinError(ctypes.get_last_error())}')
        self.path = destination

    def __exit__(self, *exc):
        if self.handle is not None:
            self.api.CloseHandle(self.handle)
            self.handle = None


class GenerationGuard:
    def __init__(self, folder, files):
        self.folder, self.files = Path(folder), files

    def __enter__(self):
        self.stack = ExitStack()
        self.handles, self.data = {}, {}
        try:
            for name in self.files:
                path = self.folder/name
                if path.exists():
                    handle = self.stack.enter_context(ProtectedFile(path))
                    self.handles[name] = handle
                    self.data[name] = handle.read()
                else:
                    self.data[name] = None
            self.hashes = {name: digest(data) if data is not None else None for name,data in self.data.items()}
            return self
        except BaseException:
            self.stack.close()
            raise

    def __exit__(self, *exc):
        self.stack.close()

    def publish(self, stage, fail_at=None):
        """Keep old AND new files protected throughout publication and rollback."""
        stage = Path(stage)
        retired = stage/'previous'
        retired.mkdir()
        moved_old, installed = [], []
        with ExitStack() as stack:
            new = {name: stack.enter_context(ProtectedFile(stage/name)) for name in self.files}
            try:
                for name in self.files:
                    if name in self.handles:
                        self.handles[name].rename(retired/name)
                        moved_old.append(name)
                    # Even in the name-vacancy interval an external writer cannot
                    # have its bytes overwritten: an existing destination fails.
                    new[name].rename(self.folder/name)
                    installed.append(name)
                    if fail_at == 'after_first_replace':
                        raise ProfileError('Injected safe-write failure during commit')
            except BaseException as original:
                blocked = []
                for name in reversed(installed):
                    try:
                        new[name].rename(stage/name)
                    except Exception as exc:
                        blocked.append(str(exc))
                for name in reversed(moved_old):
                    try:
                        self.handles[name].rename(self.folder/name)
                    except Exception as exc:
                        # Never overwrite an external file, including on rollback.
                        blocked.append(str(exc))
                if blocked:
                    raise ProfileError('Commit stopped; external destinations preserved; verified previous generation is in backup. '
                                       + '; '.join(blocked)) from original
                raise
