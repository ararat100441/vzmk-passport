"""Durable append-only journal, immutable backups and Windows file guards."""
from contextlib import contextmanager
from datetime import datetime, timezone
import hashlib
import json
import os
from pathlib import Path
import re
from uuid import uuid4

ROOT = Path(__file__).resolve().parents[1]


class Stop(ValueError):
    """Fail closed; do not retry a write on an uncertain operation."""


def now():
    return datetime.now(timezone.utc).isoformat()


def digest(value):
    return hashlib.sha256(json.dumps(value, ensure_ascii=False, sort_keys=True,
                                     separators=(',', ':')).encode('utf8')).hexdigest()


def sha(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def read(path):
    try:
        return json.loads(Path(path).read_text(encoding='utf-8-sig'))
    except (OSError, ValueError) as exc:
        raise Stop('UNREADABLE_JSON: ' + str(path)) from exc


def no_links(path):
    p = Path(path).absolute()
    for node in [p, *p.parents]:
        if node.is_symlink() or node.is_junction():
            raise Stop('LINK_PATH_FORBIDDEN: ' + str(node))
    if p.is_file() and p.stat().st_nlink != 1:
        raise Stop('HARDLINK_FORBIDDEN: ' + str(p))
    return p.resolve()


def isolated_parent(path):
    p = no_links(path)
    relative = p.relative_to(ROOT/'validation') if p.is_relative_to(ROOT/'validation') else None
    if not relative or len(relative.parts) != 2 or not relative.parts[0].startswith('stage8_') or relative.parts[1] != 'sandboxes':
        raise Stop('ISOLATED_STAGE8_DIRECTORY_REQUIRED')
    return p


def sandbox(path):
    p = no_links(path)
    isolated_parent(p.parent)
    if not re.fullmatch(r'[0-9a-f]{32}', p.name):
        raise Stop('SANDBOX_ID_INVALID')
    meta = read(p/'SANDBOX.json')
    if (meta.get('schema') != 'STAGE8_SANDBOX_V1' or meta.get('sandbox_id') != p.name
            or meta.get('mode') != 'ISOLATED_TEST' or meta.get('official_issue') is not False):
        raise Stop('SANDBOX_AUTHORITY_INVALID')
    if meta.get('seal') != digest({k:v for k,v in meta.items() if k != 'seal'}):
        raise Stop('SANDBOX_METADATA_DAMAGED')
    for sub in ('registry', 'journal', 'backups', 'transactions', 'inputs'):
        no_links(p/sub)
    no_links(p/'registry/УЗК_TEST.xlsm')
    return p, meta


def within(path, root):
    p = no_links(path)
    if not p.is_relative_to(root) or p == root:
        raise Stop('INPUT_OUTSIDE_SANDBOX: ' + str(p))
    return p


def exclusive_bytes(path, data):
    """Flush content then publish without replacing any previous journal entry."""
    p = Path(path)
    temp = p.with_name(p.name+'.'+uuid4().hex+'.tmp')
    with temp.open('xb') as f:
        f.write(data)
        f.flush()
        os.fsync(f.fileno())
    # Windows rename fails if destination exists. Temporary crash leftovers are
    # never treated as a durable journal entry; recovery reads only the prefix.
    os.rename(temp, p)


def write(path, value):
    exclusive_bytes(path, json.dumps(value, ensure_ascii=False, indent=2).encode('utf8'))


@contextmanager
def lock(root):
    import win32file
    import win32con
    root, _ = sandbox(root)
    no_links(root/'.stage8.lock')
    try:
        h = win32file.CreateFile(str(root/'.stage8.lock'), win32con.GENERIC_READ | win32con.GENERIC_WRITE,
                                0, None, win32con.OPEN_ALWAYS, 0, None)
    except Exception as exc:
        raise Stop('PARALLEL_OPERATION_OR_LOCKED_SANDBOX') from exc
    try:
        yield
    finally:
        h.Close()


@contextmanager
def frozen(paths, replaceable=None):
    """Deny external writes; registry permits only our atomic file replacement."""
    import win32file
    import win32con
    handles = []
    try:
        for path in sorted(set(map(str, paths))):
            p = no_links(path)
            share = win32con.FILE_SHARE_READ
            if replaceable is not None and p == Path(replaceable).resolve():
                share |= win32con.FILE_SHARE_DELETE
            try:
                h = win32file.CreateFile(str(p), win32con.GENERIC_READ, share,
                                        None, win32con.OPEN_EXISTING, 0, None)
            except Exception as exc:
                raise Stop('INPUT_LOCKED: '+str(p)) from exc
            handles.append(h)
        yield
    finally:
        for h in reversed(handles):
            h.Close()


def history(root):
    """Validate sequence, hash chain and legal transitions before every action."""
    directory = Path(root)/'journal'
    events, previous, active, closed, expected = [], None, None, {}, read(Path(root)/'SANDBOX.json')['initial_registry_sha256']
    files = sorted(directory.glob('*.json'))
    transitions = {'INTENT': 'BACKUP_VERIFIED', 'BACKUP_VERIFIED': 'STAGED',
                   'STAGED': 'APPLYING', 'APPLYING': 'COMMITTED'}
    for index, file in enumerate(files):
        e = read(file)
        if (file.name != f'{index:08d}.json' or e.get('sequence') != index
                or e.get('previous') != previous or e.get('schema') != 'STAGE8_EVENT_V1'):
            raise Stop('JOURNAL_CHAIN_BROKEN')
        if e.get('seal') != digest({k:v for k,v in e.items() if k != 'seal'}):
            raise Stop('JOURNAL_EVENT_DAMAGED')
        if e.get('state') == 'INTENT':
            if active is not None or e['key'] in closed or e['before_sha256'] != expected:
                raise Stop('JOURNAL_INTENT_CONFLICT')
            active = e.copy()
            if (active['plan']['mode'] != 'ISOLATED_TEST' or active['plan']['official_issue'] is not False
                    or active['preview_seal'] != digest(active['plan']) or active.get('official_issue') is not False):
                raise Stop('JOURNAL_PLAN_DAMAGED')
        else:
            if active is None or any(e.get(k) != active.get(k) for k in ('key', 'operation_id')):
                raise Stop('JOURNAL_OPERATION_MISMATCH')
            state = e.get('state')
            if state not in (transitions.get(active['state']), 'ABORTED'):
                raise Stop('JOURNAL_TRANSITION_INVALID')
            if state == 'ABORTED' and active['state'] == 'COMMITTED':
                raise Stop('JOURNAL_TRANSITION_INVALID')
            active.update(e)
            if state == 'COMMITTED':
                expected = active['after_sha256']
                closed[active['key']] = active.copy()
                active = None
            elif state == 'ABORTED':
                active = None
        events.append(e)
        previous = sha(file)
    return {'events': events, 'active': active, 'closed': closed, 'expected_sha256': expected}


def append(root, state, operation, **fields):
    h = history(root)
    previous = sha(Path(root)/'journal'/f'{len(h["events"])-1:08d}.json') if h['events'] else None
    event = {'schema': 'STAGE8_EVENT_V1', 'sequence': len(h['events']), 'previous': previous,
             'time': now(), 'state': state, 'operation_id': operation['operation_id'], 'key': operation['key']} | fields
    event['seal'] = digest(event)
    write(Path(root)/'journal'/f'{event["sequence"]:08d}.json', event)
    history(root)
    return event


def backup(root, operation, registry, expected):
    folder = Path(root)/'backups'/operation['operation_id']
    folder.mkdir(exist_ok=False)
    output = folder/'УЗК_TEST.xlsm'
    exclusive_bytes(output, Path(registry).read_bytes())
    if sha(registry) != expected or sha(output) != expected:
        raise Stop('BACKUP_SHA256_MISMATCH')
    write(folder/'SHA256.json', {'path': str(output), 'sha256': expected, 'verified': True, 'time': now()})
    return {'path': str(output), 'sha256': expected, 'metadata': str(folder/'SHA256.json'),
            'metadata_sha256': sha(folder/'SHA256.json')}


def verify_backup(reference, root):
    p = within(reference['path'], root/'backups')
    m = read(within(reference['metadata'], root/'backups'))
    if (sha(reference['metadata']) != reference['metadata_sha256'] or sha(p) != reference['sha256']
            or m.get('sha256') != reference['sha256'] or m.get('verified') is not True or m.get('path') != str(p)):
        raise Stop('BACKUP_DAMAGED')


def replace(stage, registry):
    import ctypes
    import win32file
    import win32con
    import win32security
    # Same-volume atomic rename by handle, retaining BOTH deny-write guards.
    # FileRenameInfoEx POSIX semantics permit our old read handle to survive;
    # the stage's read guard continues protecting the new target after rename.
    descriptor = win32security.GetFileSecurity(str(registry),win32security.DACL_SECURITY_INFORMATION)
    win32security.SetFileSecurity(str(stage),win32security.DACL_SECURITY_INFORMATION,descriptor)
    name = str(Path(registry).resolve())
    class RenameInfo(ctypes.Structure):
        _fields_ = [('Flags',ctypes.c_uint32),('RootDirectory',ctypes.c_void_p),
                    ('FileNameLength',ctypes.c_uint32),('FileName',ctypes.c_wchar*len(name))]
    info = RenameInfo(0x1 | 0x2,None,len(name.encode('utf-16-le')),name)
    fn = ctypes.WinDLL('kernel32',use_last_error=True).SetFileInformationByHandle
    fn.argtypes = [ctypes.c_void_p,ctypes.c_int,ctypes.c_void_p,ctypes.c_uint32]
    fn.restype = ctypes.c_int
    h = win32file.CreateFile(str(stage),win32con.GENERIC_READ | win32con.DELETE,
                            win32con.FILE_SHARE_READ | win32con.FILE_SHARE_DELETE,
                            None,win32con.OPEN_EXISTING,0,None)
    try:
        if not fn(int(h),22,ctypes.byref(info),ctypes.sizeof(info)):
            raise ctypes.WinError(ctypes.get_last_error())
    finally:
        h.Close()


def flush_file(path):
    with Path(path).open('r+b') as f:
        f.flush()
        os.fsync(f.fileno())
