"""Private, pinned local workspace for Office, independent of source path length."""
from contextlib import contextmanager
import os
from pathlib import Path
import sys
from uuid import uuid4
from . import storage as S

# A conservative application policy, not a Windows/NTFS universal limit.
BASE = Path(os.environ.get('LOCALAPPDATA', ''))
MAX_OFFICE_PATH = 180


def _remove_owned(path, identity):
    import ctypes
    import win32file, win32con
    h = win32file.CreateFile(str(path), win32con.GENERIC_READ | win32con.DELETE,
        win32con.FILE_SHARE_READ, None, win32con.OPEN_EXISTING,
        win32con.FILE_FLAG_BACKUP_SEMANTICS | win32file.FILE_FLAG_OPEN_REPARSE_POINT, None)
    try:
        if S._identity(h) != identity: raise S.Stop('NATIVE_CLEANUP_IDENTITY_CHANGED')
        flag = ctypes.c_int(1)
        fn = ctypes.WinDLL('kernel32', use_last_error=True).SetFileInformationByHandle
        fn.argtypes = [ctypes.c_void_p,ctypes.c_int,ctypes.c_void_p,ctypes.c_uint32]
        fn.restype = ctypes.c_int
        if not fn(int(h),4,ctypes.byref(flag),ctypes.sizeof(flag)):
            raise ctypes.WinError(ctypes.get_last_error())
    finally: h.Close()


class Workspace:
    def __init__(self, path):
        self.path, self.owned, self.active = path, {}, True

    def __truediv__(self, name): return self.path/name

    def track(self, path):
        if path.parent != self.path: raise S.Stop('NATIVE_WORKSPACE_PATH_MISMATCH')
        self.owned[path] = S.file_identity(path)

    def remove(self, path):
        _remove_owned(path, self.owned[path])
        del self.owned[path]


@contextmanager
def workspace():
    import win32api, win32con, win32file, win32security, pywintypes
    folder = BASE / ('s8n_' + uuid4().hex)
    if not BASE.is_absolute() or not BASE.is_dir() or len(str(folder/'stage.xlsm')) > MAX_OFFICE_PATH:
        raise S.Stop('NATIVE_SHORT_WORKSPACE_UNAVAILABLE')
    S.no_links(BASE)
    token = win32security.OpenProcessToken(win32api.GetCurrentProcess(), win32con.TOKEN_QUERY)
    try: sid = win32security.ConvertSidToStringSid(win32security.GetTokenInformation(token, win32security.TokenUser)[0])
    finally: token.Close()
    sd = win32security.ConvertStringSecurityDescriptorToSecurityDescriptor(
        'D:P(A;OICI;FA;;;SY)(A;OICI;FA;;;' + sid + ')', 1)
    sa = pywintypes.SECURITY_ATTRIBUTES(); sa.SECURITY_DESCRIPTOR = sd
    try:
        # ACL applies at creation, with no shared predictable temporary file.
        with S.namespace_guard(BASE):
            win32file.CreateDirectory(str(folder), sa)
            own = Workspace(folder)
            dh = win32file.CreateFile(str(folder),win32con.GENERIC_READ,
                win32con.FILE_SHARE_READ | win32con.FILE_SHARE_WRITE | win32con.FILE_SHARE_DELETE,
                None,win32con.OPEN_EXISTING,win32con.FILE_FLAG_BACKUP_SEMANTICS,None)
            try: directory_identity = S._identity(dh)
            finally: dh.Close()
            try:
                with S.namespace_guard(folder):
                    try:
                        yield own
                    finally:
                        original = sys.exc_info()[1]
                        try:
                            # Only owned identities in this unique workspace.
                            for p in folder.iterdir():
                                if p not in own.owned or not p.is_file():
                                    raise S.Stop('NATIVE_WORKSPACE_UNEXPECTED_LEFTOVER: '+str(p))
                                own.remove(p)
                        except Exception as exc:
                            if original is None: raise S.Stop('NATIVE_WORKSPACE_CLEANUP_FAILED') from exc
                            original.add_note('NATIVE_WORKSPACE_CLEANUP_FAILED: '+str(exc))
            finally:
                own.active = False
                original = sys.exc_info()[1]
                try:
                    # Release the namespace pin before requesting DELETE, then
                    # compare identity on the deleting handle, without a gap
                    # between that comparison and disposition of that object.
                    _remove_owned(folder, directory_identity)
                except Exception as exc:
                    if original is None: raise S.Stop('NATIVE_WORKSPACE_DIRECTORY_CLEANUP_FAILED') from exc
                    original.add_note('NATIVE_WORKSPACE_DIRECTORY_CLEANUP_FAILED: '+str(exc))
    except S.Stop:
        raise
    except (OSError, pywintypes.error) as exc:
        raise S.Stop('NATIVE_SHORT_WORKSPACE_FAILURE: '+str(exc)) from exc


@contextmanager
def exact_copy(path, folder, *, guard=None):
    """Hold source identity/content and copy identity/content through Excel.

    A caller's existing write/delete guard is retained, never reacquired with
    weaker sharing. In that case read the exact source through its own handle.
    """
    from contextlib import ExitStack
    import hashlib
    source = Path(path).absolute()
    if not isinstance(folder, Workspace) or not folder.active:
        raise S.Stop('NATIVE_UNPROTECTED_VERIFICATION_WORKSPACE')
    copy = folder/'v.xlsm'
    with ExitStack() as stack:
        if guard is None:
            stack.enter_context(S.namespace_guard(source.parent))
            stack.enter_context(S.frozen([source]))
            identity = S.file_identity(source)
            data = source.read_bytes()
        else:
            if source != guard.path: raise S.Stop('NATIVE_SOURCE_GUARD_PATH_MISMATCH')
            identity = guard.identity
            data = guard.read_bytes()
        before = hashlib.sha256(data).hexdigest()
        if guard is not None: guard.check(identity, before)
        with copy.open('xb') as f:
            f.write(data); f.flush(); os.fsync(f.fileno())
        folder.track(copy)
        copy_identity = S.file_identity(copy)
        stack.enter_context(S.frozen([copy]))
        if S.sha(copy) != before or S.file_identity(copy) != copy_identity:
            raise S.Stop('NATIVE_SHORT_COPY_CHANGED')
        binding = {'source_path':str(source),'source_identity':identity,'source_sha256':before,
                   'verification_path':str(copy),'verification_identity':copy_identity,
                   'verification_sha256':before,'exact_bytes':True}
        yield copy, binding
        if S.sha(copy) != before or S.file_identity(copy) != copy_identity:
            raise S.Stop('NATIVE_SHORT_COPY_CHANGED')
        if guard is not None: guard.check(identity,before)
        elif S.sha(source) != before or S.file_identity(source) != identity:
            raise S.Stop('NATIVE_SOURCE_CHANGED')
        binding['source_and_copy_rechecked'] = True
    folder.remove(copy)
