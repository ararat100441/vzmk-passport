"""One-shot guarded restoration of exactly the accepted three-line M1 patch."""
import hashlib
import json
from pathlib import Path
import win32con
import win32file
import pywintypes

RUN = Path(__file__).resolve().parent
ROOT = RUN.parents[1]
TARGET = ROOT / 'passport_app/vom.py'


def sha(data):
    return hashlib.sha256(data).hexdigest()


def main():
    pre = json.loads((RUN / 'PRECHECK.json').read_text(encoding='utf-8'))
    tests = json.loads((RUN / 'TESTS_BEFORE.json').read_text(encoding='utf-8'))
    assert tests['stage6_smoke'] and tests['selected_vom_pass'] and tests['m1_passed'] == 2
    baseline = json.loads((RUN / 'PRODUCT_BEFORE.json').read_text(encoding='utf-8'))
    assert all(sha(Path(p).read_bytes()) == h for p, h in baseline.items()), 'Product drift before mutation'
    accepted = (RUN / 'vom.accepted.py').read_bytes()
    assert sha(accepted) == pre['accepted_sha256']
    # OPEN_EXISTING with share=0: fail before writing if any other handle is open;
    # prevent competing read/write/delete/replace handles throughout this patch.
    handle = win32file.CreateFile(str(TARGET), win32con.GENERIC_READ | win32con.GENERIC_WRITE,
                                 0, None, win32con.OPEN_EXISTING, 0, None)
    try:
        stat = TARGET.stat()
        _, current = win32file.ReadFile(handle, 1024 * 1024)
        assert sha(current) == pre['current_sha256'] == pre['expected_audit_sha'], 'Unexpected SHA drift'
        assert stat.st_mtime_ns == pre['mtime_ns'] and stat.st_size == pre['size'], 'Unexpected metadata drift'
        original = b"            trimmed = (name or '').strip()\n            if trimmed and trimmed not in names:\n                names.append(trimmed)"
        restored = b"            original_name = name or ''\n            if original_name.strip() and original_name not in names:\n                names.append(original_name)"
        assert current.count(original) == 1 and current.replace(original, restored) == accepted
        try:
            competing = win32file.CreateFile(str(TARGET), win32con.GENERIC_WRITE,
                                             win32con.FILE_SHARE_READ | win32con.FILE_SHARE_WRITE,
                                             None, win32con.OPEN_EXISTING, 0, None)
        except pywintypes.error as exc:
            assert exc.winerror == 32, exc
            competing_error = exc.winerror
        else:
            competing.Close()
            raise RuntimeError('Exclusive concurrency guard not enforced; no mutation permitted')
        win32file.SetFilePointer(handle, 0, win32con.FILE_BEGIN)
        _, written = win32file.WriteFile(handle, accepted)
        assert written == len(accepted)
        win32file.SetEndOfFile(handle)
        win32file.FlushFileBuffers(handle)
        win32file.SetFilePointer(handle, 0, win32con.FILE_BEGIN)
        _, final = win32file.ReadFile(handle, 1024 * 1024)
        assert final == accepted and sha(final) == pre['accepted_sha256']
    finally:
        handle.Close()
    report = {'pass': True, 'target': str(TARGET), 'before_sha256': sha(current),
              'after_sha256': sha(final), 'before_mtime_ns': pre['mtime_ns'],
              'after_mtime_ns': TARGET.stat().st_mtime_ns,
              'exclusive_handle': 'GENERIC_READ|GENERIC_WRITE, share=0, OPEN_EXISTING',
              'competing_write_probe_winerror': competing_error,
              'exact_accepted_bytes': True, 'patch': 'Three Connections lines only; other bytes identical'}
    (RUN / 'CONCURRENCY_AND_RESTORE.json').write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding='utf-8')
    (RUN / 'vom.after.py').write_bytes(final)
    print(json.dumps(report, ensure_ascii=False))


if __name__ == '__main__':
    main()
