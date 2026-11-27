"""Preserve operation evidence and redundant Excel replacement copies."""
import hashlib
import json
import os
import shutil
from pathlib import Path

from .storage import Stop, mutex, post_commit_path, read, safe, sha


EVIDENCE_FILES = ('plan.json', 'prepared.json', 'verified.json', 'confirmation.json', 'release.json')


def operation_evidence_writes(stage, run):
    for name in EVIDENCE_FILES:
        source = safe(Path(run) / name)
        yield Path(stage) / 'evidence' / name, source.stat().st_size


def require_disk_space(writes):
    """Count required bytes together when several destinations share one volume."""
    volumes = {}
    for destination, size in writes:
        folder = safe(destination).parent
        while not folder.is_dir() and folder != folder.parent:
            folder = folder.parent
        try:
            device = folder.stat().st_dev
        except OSError as exc:
            raise Stop(f'Не удалось проверить диск назначения {destination}: {exc}') from exc
        anchor, total = volumes.get(device, (folder, 1024 * 1024))  # Journal and filesystem metadata.
        volumes[device] = anchor, total + size
    for folder, required in volumes.values():
        try:
            free = shutil.disk_usage(folder).free
        except OSError as exc:
            raise Stop(f'Не удалось определить свободное место на диске {folder}: {exc}') from exc
        if free < required:
            raise Stop(f'Недостаточно места на диске {folder}. Требуется {required:,} байт, доступно {free:,} байт.')


def retain_operation_evidence(stage, run, plan, *, prepared, confirmation, released):
    """Preserve exact bytes for the files checked by post-commit and refunds."""
    expected = {'plan.json': plan, 'prepared.json': prepared,
                'confirmation.json': confirmation, 'release.json': released}
    payloads, hashes = {}, {}
    for name in EVIDENCE_FILES:
        source = safe(Path(run) / name)
        try:
            payloads[name] = source.read_bytes()
            value = json.loads(payloads[name].decode('utf-8'))
        except (OSError, ValueError) as exc:
            raise Stop(f'Не удалось сохранить доказательство операции {source}: {exc}') from exc
        hashes[name] = hashlib.sha256(payloads[name]).hexdigest()
        if name in expected and value != expected[name]:
            raise Stop('Доказательство операции изменилось перед записью: ' + str(source))
    if confirmation.get('verified_sha256') != hashes['verified.json']:
        raise Stop('Проверка выпуска изменилась перед записью: ' + str(Path(run) / 'verified.json'))
    folder = safe(Path(stage) / 'evidence')
    folder.mkdir()
    for name, payload in payloads.items():
        destination = safe(folder / name)
        with destination.open('xb') as stream:
            stream.write(payload)
            stream.flush()
            os.fsync(stream.fileno())
        if sha(safe(destination)) != hashes[name]:
            raise Stop('Не совпал SHA-256 архива операции: ' + str(destination))
    return hashes


def operation_evidence(journal, state, *, names=EVIDENCE_FILES, legacy_post_commit=False):
    """Read the sealed archive, or use the original paths of historical journals."""
    if set(names) - set(EVIDENCE_FILES):
        raise ValueError('Unknown operation evidence filename')
    if 'evidence' not in state:
        plan = Path(state['plan'])
        check_plan = post_commit_path if legacy_post_commit else safe
        return {name: check_plan(plan) if name == 'plan.json' else safe(plan.parent / name)
                for name in names}
    hashes = state['evidence']
    if not isinstance(hashes, dict) or set(hashes) != set(EVIDENCE_FILES):
        raise Stop('Повреждена опись доказательств операции: ' + str(journal))
    folder = safe(Path(journal).parent / 'evidence')
    paths = {}
    for name in names:
        path = safe(folder / name)
        try:
            current = sha(path)
        except OSError as exc:
            raise Stop(f'Доказательство операции отсутствует или недоступно {path}: {exc}') from exc
        if current != hashes[name]:
            raise Stop('Не совпал SHA-256 доказательства операции: ' + str(path))
        paths[name] = path
    return paths


def _record(path, data):
    # Do not replace an existing audit record. A durable intent precedes rename.
    with path.open('x', encoding='utf-8') as stream:
        json.dump(data, stream, ensure_ascii=False, indent=2)
        stream.flush()
        os.fsync(stream.fileno())


def _rename_and_verify(source, destination, expected):
    import win32con
    import win32file
    from pywintypes import error as Win32Error

    # Deny writers while validating the bytes and moving the same file.
    try:
        handle = win32file.CreateFile(
            str(source), win32con.GENERIC_READ,
            win32con.FILE_SHARE_READ | win32con.FILE_SHARE_DELETE,
            None, win32con.OPEN_EXISTING, 0x00200000, None)  # OPEN_REPARSE_POINT
    except Win32Error as exc:
        raise Stop(f'Служебная копия занята или недоступна: {source}; {exc}') from exc
    try:
        if destination.exists() or sha(source) != expected:
            raise Stop('Файл изменился или место в архиве уже занято: ' + str(source))
        if source.stat().st_dev != destination.parent.stat().st_dev:
            raise Stop('Источник и архив находятся на разных томах: ' + str(source))
        source.rename(destination)  # Windows fails if destination exists.
        try:
            if sha(destination) != expected or source.exists():
                raise Stop('SHA-256 после переноса не подтверждён: ' + str(destination))
        except (OSError, Stop) as exc:
            try:
                destination.rename(source)  # No overwrite on rollback either.
            except OSError as restore_error:
                raise Stop(f'Не подтверждён перенос {source} → {destination}; '
                           f'проверьте оба пути: {restore_error}') from exc
            raise
    finally:
        handle.Close()


def move_verified(source, destination, retained_backup, expected):
    """Same-volume, no-overwrite rename, with the independent backup retained."""
    source, destination, backup = map(safe, (source, destination, retained_backup))
    if len({source, destination, backup}) != 3 or sha(backup) != expected:
        raise Stop('Не подтверждена независимая резервная копия: ' + str(backup))
    if os.name != 'nt':
        raise Stop('Архивирование требует переименования Windows без перезаписи.')
    destination.parent.mkdir(parents=True, exist_ok=True)
    intent = safe(destination.parent / 'move-intent.json')
    verified = safe(destination.parent / 'move-verified.json')
    record = {'source': str(source), 'destination': str(destination),
              'retained_backup': str(backup), 'sha256': expected}
    if intent.exists():
        if read(intent) != record:
            raise Stop('Другое намерение переноса уже записано: ' + str(intent))
    else:
        if destination.exists() or verified.exists() or sha(source) != expected:
            raise Stop('Не подтверждён исходный файл или архив уже занят: ' + str(source))
        _record(intent, record)
    if verified.exists():
        if read(verified) == record and not source.exists() and sha(destination) == expected:
            return record
        raise Stop('Ранее архивированная копия изменилась: ' + str(destination))
    if not source.exists():
        # A previous process may have exited between rename and verification.
        if sha(destination) != expected:
            raise Stop('Не подтверждён незавершённый перенос: ' + str(destination))
    else:
        _rename_and_verify(source, destination, expected)
    _record(verified, record)
    return record


def path_is_referenced(value, path):
    if isinstance(value, dict):
        return any(path_is_referenced(part, path) for item in value.items() for part in item)
    if isinstance(value, list):
        return any(path_is_referenced(item, path) for item in value)
    # Include relative path references and strings with a path embedded in them.
    return isinstance(value, str) and path.name.casefold() in value.casefold()


def replacement_path(path, operation_id):
    path = Path(path)
    return path.with_name('.vzmk-replaced-' + operation_id + path.suffix)


def replacement_paths(items, operation_id):
    # Historical names stay valid for every layout accepted by older builds.
    paths = [replacement_path(item['path'], operation_id) for item in items]
    return [replacement_path(item['path'], operation_id + '-' + item['role'])
            if paths.count(path) > 1 else path for item, path in zip(items, paths)]


def archive_completed(journal):
    """Only exact, unreferenced duplicates of sealed coordinated backups."""
    from .core import journal_read

    journal = safe(journal)
    state = journal_read(journal)
    stage = journal.parent
    registry = Path(state['registry'])
    # Serialize the reference check with other operations for this registry.
    # Failure to take this optional lock leaves the copy where it was.
    with mutex(registry, post_commit=True):
        if (state['phase'] != 'COMPLETE' or stage.name != state['id']
                or stage.parent != registry.parent / '.vzmk-operations'
                or journal_read(journal) != state):
            raise Stop('Архивация допустима только для подтверждённого COMPLETE.')
        return _archive_unreferenced(journal, state)


def _archive_unreferenced(journal, state):
    from .core import journal_read

    stage = journal.parent
    plan, target = Path(state['plan']), Path(state['target'])
    names = EVIDENCE_FILES if 'evidence' in state else ('plan.json',)
    evidence = {journal, target / 'ISSUE.json', *operation_evidence(journal, state, names=names).values()}
    evidence.update(stage.glob('*.json'))
    evidence.update(plan.parent.glob('*.json'))
    terminal = {'COMPLETE', 'COMPENSATED', 'ABORTED', 'ROLLED_BACK'}
    for sibling in stage.parent.glob('*/journal.json'):
        if sibling == journal:
            continue
        other = journal_read(safe(sibling))
        if other['phase'] not in terminal:
            other_plan = safe(other['plan'])
            names = EVIDENCE_FILES if 'evidence' in other else ('plan.json',)
            evidence.add(sibling)
            evidence.update(operation_evidence(sibling, other, names=names).values())
            evidence.update(sibling.parent.rglob('*.json'))
            evidence.update(other_plan.parent.glob('*.json'))
    references = [read(safe(path)) for path in evidence]
    result = {'moved': [], 'retained': []}
    items = state['participants']
    roles = {'registry', 'catalog', 'uzk'}
    if len(items) != 3 or {item['role'] for item in items} != roles:
        raise Stop('Не подтверждены участники завершённой операции.')
    sources = replacement_paths(items, state['id'])
    if len(set(sources)) != len(items):
        raise Stop('Совпадают пути служебных копий участников операции.')
    for item, source in zip(items, sources):
        try:
            role = item['role']
            backup = safe(item['backup'])
            if backup != stage / ('backup-' + role + Path(item['path']).suffix):
                raise Stop('Резервная копия находится вне штатного пути: ' + str(backup))
            if any(path_is_referenced(value, source) for value in references):
                raise Stop('На копию ссылаются документы операции.')
            # A project can live on another volume. Keep the rename local there.
            archive = stage / 'service-archive'
            if source.parent.stat().st_dev != stage.stat().st_dev:
                archive = source.parent / '.vzmk-service-archive' / state['id']
            result['moved'].append(move_verified(
                source, archive / role / source.name, backup, item['original_sha256']))
        except Exception as exc:  # A housekeeping error must not skip the other books.
            result['retained'].append({'path': str(source), 'reason': str(exc)})
    return result
