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
def frozen(paths):
    """Deny both writes and namespace replacement until every handle is closed."""
    import win32file
    import win32con
    handles = []
    try:
        for path in sorted(set(map(str, paths))):
            p = no_links(path)
            try:
                h = win32file.CreateFile(str(p), win32con.GENERIC_READ, win32con.FILE_SHARE_READ,
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


def _identity(handle):
    import win32file
    info = win32file.GetFileInformationByHandle(handle)
    if info[7] != 1 or info[0] & 0x400:
        raise Stop('REGISTRY_LINK_OR_REPARSE_FORBIDDEN')
    return {'volume_serial':info[4], 'file_index':(info[8] << 32) | info[9]}


def file_identity(path):
    import win32file, win32con
    h = win32file.CreateFile(str(path), win32con.GENERIC_READ,
        win32con.FILE_SHARE_READ | win32con.FILE_SHARE_WRITE | win32con.FILE_SHARE_DELETE,
        None, win32con.OPEN_EXISTING, win32file.FILE_FLAG_OPEN_REPARSE_POINT, None)
    try:
        return _identity(h)
    finally:
        h.Close()


class RegistryGuard:
    """One read/write handle; never release or upgrade it to publish.

    Native Excel saves a private staged XLSM. Its verified bytes are copied
    through this handle. This deliberately is NOT an atomic whole-file rename:
    death mid-copy leaves APPLYING/ambiguous, with an immutable verified backup.
    Ordinary readers may see a partial copy; they must not infer CLOSED from it.
    """
    def __init__(self, path, handle):
        self.path, self.handle = path, handle
        self.identity = _identity(handle)

    def check(self, expected_identity, expected_sha):
        if self.identity != expected_identity or file_identity(self.path) != self.identity:
            raise Stop('REGISTRY_FILE_ID_CHANGED')
        if self.sha() != expected_sha:
            raise Stop('REGISTRY_CONTENT_CHANGED')

    def sha(self):
        import win32file
        win32file.SetFilePointer(self.handle, 0, 0)
        result = hashlib.sha256()
        while True:
            try:
                _, block = win32file.ReadFile(self.handle, 1024*1024)
            except Exception as exc:
                if getattr(exc, 'winerror', None) == 38:
                    break
                raise
            if not block:
                break
            result.update(block)
        return result.hexdigest()

    def publish(self, stage, expected_identity, before, after, fault=None):
        import win32file
        self.check(expected_identity, before)
        if sha(stage) != after:
            raise Stop('STAGE_CONTENT_CHANGED')
        win32file.SetFilePointer(self.handle, 0, 0)
        # Small first block provides a deterministic partial-write crash seam.
        # There is no unguarded pathname write, rename, or handle re-acquisition.
        with Path(stage).open('rb') as source:
            block = source.read(4096)
            while block:
                _, written = win32file.WriteFile(self.handle, block)
                if written != len(block):
                    raise Stop('REGISTRY_SHORT_WRITE')
                if fault:
                    fault('during_publish', self.path)
                block = source.read(1024*1024)
        win32file.SetEndOfFile(self.handle)
        win32file.FlushFileBuffers(self.handle)
        self.check(expected_identity, after)


@contextmanager
def registry_guard(path):
    """NTFS local volume only. Pin the path's ancestors and its file identity.

    No FILE_SHARE_WRITE/DELETE on the target or directory ancestors. An existing
    writer/deleter makes acquisition fail, including handles opened beforehand.
    The guard is held through validation, backup, copy, verification and COMMIT.
    """
    import win32file, win32con, win32api, pywintypes
    p = Path(path).absolute()
    handles = []
    try:
        if (win32file.GetDriveType(p.anchor) != win32con.DRIVE_FIXED
                or win32api.GetVolumeInformation(p.anchor)[4] != 'NTFS'):
            raise Stop('LOCAL_NTFS_REQUIRED')
        for directory in reversed(p.parents):
            h = win32file.CreateFile(str(directory), 0x80,  # FILE_READ_ATTRIBUTES
                win32con.FILE_SHARE_READ, None, win32con.OPEN_EXISTING,
                win32con.FILE_FLAG_BACKUP_SEMANTICS | win32file.FILE_FLAG_OPEN_REPARSE_POINT, None)
            handles.append(h)
            if win32file.GetFileInformationByHandle(h)[0] & win32con.FILE_ATTRIBUTE_REPARSE_POINT:
                raise Stop('REGISTRY_DIRECTORY_REPARSE_FORBIDDEN')
        h = win32file.CreateFile(str(p), win32con.GENERIC_READ | win32con.GENERIC_WRITE,
            win32con.FILE_SHARE_READ, None, win32con.OPEN_EXISTING,
            win32file.FILE_FLAG_OPEN_REPARSE_POINT, None)
        handles.append(h)
        no_links(p)
        yield RegistryGuard(p, h)
    except Stop:
        raise
    except (OSError, pywintypes.error) as exc:
        raise Stop('REGISTRY_GUARD_OR_IO_FAILURE: '+str(exc)) from exc
    finally:
        for h in reversed(handles):
            h.Close()


def flush_file(path):
    with Path(path).open('r+b') as f:
        f.flush()
        os.fsync(f.fileno())
