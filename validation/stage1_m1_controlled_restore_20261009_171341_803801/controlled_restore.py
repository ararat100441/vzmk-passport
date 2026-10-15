"""Restore a verified candidate under a local exclusive handle and fresh checkpoint."""
import ctypes
from ctypes import wintypes
import difflib
import hashlib
import json
import os
from pathlib import Path
import time
import win32con
import win32file
import pywintypes

RUN = Path(__file__).resolve().parent
ROOT = RUN.parents[1]
TARGET = ROOT / 'passport_app/vom.py'


def digest(data):
    return hashlib.sha256(data).hexdigest()


def write(name, value):
    (RUN / name).write_text(json.dumps(value, ensure_ascii=False, indent=2), encoding='utf-8')


def native_snapshot(handle):
    kernel = ctypes.WinDLL('kernel32', use_last_error=True)
    kernel.GetFileTime.argtypes = [wintypes.HANDLE, ctypes.c_void_p, ctypes.c_void_p,
                                   ctypes.POINTER(wintypes.FILETIME)]
    modified = wintypes.FILETIME()
    if not kernel.GetFileTime(int(handle), None, None, ctypes.byref(modified)):
        raise ctypes.WinError(ctypes.get_last_error())
    filetime = (modified.dwHighDateTime << 32) | modified.dwLowDateTime
    win32file.SetFilePointer(handle, 0, win32con.FILE_BEGIN)
    _, data = win32file.ReadFile(handle, 1024 * 1024)
    return {'sha256': digest(data), 'size': len(data),
            'mtime_ns': (filetime - 116444736000000000) * 100}, data


def main():
    pre = json.loads((RUN / 'PRECHECK.json').read_text(encoding='utf-8'))
    expected = {'sha256': pre['current_sha256'], 'size': pre['size'], 'mtime_ns': pre['mtime_ns']}
    candidate = (RUN / 'vom.candidate.py').read_bytes()
    assert digest(candidate) == pre['accepted_sha256'] == 'c63bbd362763ba99f85de1610c01622ac230f3303b7a1e297345d02259ea954f'
    assert candidate == (RUN / 'vom.accepted.py').read_bytes()
    tests = json.loads((RUN / 'TESTS_BEFORE.json').read_text(encoding='utf-8'))
    assert tests['stage6_smoke'] and tests['selected_vom_pass'] and tests['m1_passed'] == 2
    # No product module may have drifted since this new checkpoint.
    baseline = json.loads((RUN / 'PRODUCT_BEFORE.json').read_text(encoding='utf-8'))
    assert all(digest(Path(p).read_bytes()) == h for p, h in baseline.items()), 'PRODUCT_DRIFT'
    assert TARGET.stat().st_mtime_ns == expected['mtime_ns'], 'PRE_LOCK_MTIME_DRIFT'
    handle = win32file.CreateFile(str(TARGET), win32con.GENERIC_READ | win32con.GENERIC_WRITE,
                                 0, None, win32con.OPEN_EXISTING, 0, None)
    samples = []
    try:
        for delay in (0, 0.5, 0.5):
            if delay:
                time.sleep(delay)
            snapshot, raw = native_snapshot(handle)
            samples.append(snapshot)
            if snapshot != expected or raw != (RUN / 'vom.before.py').read_bytes():
                write('NEW_DRIFT_STOP.json', {'expected': expected, 'observed': snapshot,
                                              'status': 'STOP_BEFORE_WRITE', 'cause': 'NOT_ESTABLISHED'})
                raise RuntimeError('FRESH_CHECKPOINT_DRIFT; product not written')
        # Prove that another writer cannot open the file while this guard is held.
        try:
            competing = win32file.CreateFile(str(TARGET), win32con.GENERIC_WRITE,
                                             win32con.FILE_SHARE_READ | win32con.FILE_SHARE_WRITE,
                                             None, win32con.OPEN_EXISTING, 0, None)
        except pywintypes.error as exc:
            assert exc.winerror == 32, exc
            competing_error = exc.winerror
        else:
            competing.Close()
            raise RuntimeError('CONCURRENCY_GUARD_NOT_ENFORCED')
        rollback = RUN / 'vom.rollback.py'
        with rollback.open('xb') as stream:
            stream.write(raw)
            stream.flush()
            os.fsync(stream.fileno())
        assert digest(rollback.read_bytes()) == expected['sha256']
        # Final check after rollback persistence, immediately before product write.
        final_check, current = native_snapshot(handle)
        assert final_check == expected and current == raw, 'FINAL_CHECKPOINT_DRIFT'
        write('PREWRITE_GUARD.json', {'pass': True, 'expected': expected, 'samples': samples,
                                    'final_check': final_check, 'exclusive_share_mode': 0,
                                    'competing_writer_winerror': competing_error,
                                    'rollback_sha256': digest(rollback.read_bytes())})
        win32file.SetFilePointer(handle, 0, win32con.FILE_BEGIN)
        _, written = win32file.WriteFile(handle, candidate)
        assert written == len(candidate)
        win32file.SetEndOfFile(handle)
        win32file.FlushFileBuffers(handle)
        after, actual = native_snapshot(handle)
        assert actual == candidate and after['sha256'] == pre['accepted_sha256']
    finally:
        handle.Close()
    assert TARGET.read_bytes() == candidate
    diff = ''.join(difflib.unified_diff(raw.decode('utf-8').splitlines(True),
                                       actual.decode('utf-8').splitlines(True),
                                       fromfile='checkpoint/vom.py', tofile='restored/vom.py'))
    (RUN / 'EXACT_PRODUCT_DIFF.diff').write_text(diff, encoding='utf-8')
    (RUN / 'vom.after.py').write_bytes(actual)
    write('RESTORE_RESULT.json', {'patch_applied': True, 'before': expected, 'after': after,
        'exact_accepted_bytes': True, 'rollback': str(rollback), 'local_concurrency_guard': 'HELD_THROUGH_WRITE_AND_FLUSH',
        'write_method': 'Persisted rollback; write verified candidate through exclusive OPEN_EXISTING handle; flush; reread exact bytes. No atomic-rename claim.',
        'global_guard_changes': False, 'previous_mtime_event': 'UNEXPLAINED; NOT ATTRIBUTED'})
    print(json.dumps({'patch_applied': True, 'after_sha256': after['sha256'], 'exclusive_guard': True}))


if __name__ == '__main__':
    main()
