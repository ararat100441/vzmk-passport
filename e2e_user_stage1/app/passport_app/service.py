import hashlib
import json
import os
import uuid
from contextlib import ExitStack
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path

from .appendix import find_template, generate_appendix, inspect_template
from .config import PassportError
from .checks import verify_saved
from .excel import excel_session, locked_file, sha256, atomic_replace
from .registry import read_registry, select_project, append_row
from .vom import read_vom


@dataclass(frozen=True)
class Plan:
    registry: object
    project: object
    vom: object
    template: Path
    template_digest: str
    passport_root: Path
    folder: Path
    day: date

    def summary(self):
        def fmt(value, places):
            return f'{value:,.{places}f}'.replace(',', ' ').replace('.', ',')
        return '\n'.join([
            f'Проект: {self.project.name}', f'Первый паспорт: {self.project.first_number} (строка {self.project.first_row})',
            f'Последний паспорт: {self.project.last_number}', f'Новый паспорт: {self.project.next_number}',
            f'Дата окончания: {self.day:%d.%m.%Y}', f'Строк ВОМ: {len(self.vom.rows)}',
            f'Изделий: {self.vom.quantity:.0f}', f'Масса: {fmt(self.vom.mass, 2)} кг',
            f'Площадь: {fmt(self.vom.area, 3)} м²', f'Соединения: {self.vom.connections}',
            f'Папка: {self.folder}', 'Ошибок: 0',
        ])


def writable_directory(path):
    """Checks NTFS add-file/add-subdirectory permissions without creating a file."""
    import win32con
    import win32file
    try:
        handle = win32file.CreateFile(str(path), 0x0002 | 0x0004,
                                     win32con.FILE_SHARE_READ | win32con.FILE_SHARE_WRITE | win32con.FILE_SHARE_DELETE,
                                     None, win32con.OPEN_EXISTING, win32con.FILE_FLAG_BACKUP_SEMANTICS, None)
        handle.Close()
    except Exception as exc:
        raise PassportError(f'Нет права создавать файлы и папки в {path}.') from exc


def ensure_no_pending(registry_path):
    transaction_root = registry_path.parent / '.pasport-transactions'
    if transaction_root.exists():
        for path in transaction_root.glob('*/journal.json'):
            try:
                status = json.loads(path.read_text(encoding='utf-8'))['phase']
            except Exception as exc:
                raise PassportError(f'Не удалось проверить предыдущую операцию: {path}') from exc
            if status not in {'complete', 'aborted', 'rolled_back'}:
                raise PassportError(f'Обнаружена незавершённая операция ({status}): {path}. Автоматический повтор запрещён.')


def validate(settings, project_name, vom_path, *, today=None):
    if not project_name:
        raise PassportError('Выберите проект.')
    if not vom_path:
        raise PassportError('Выберите отгрузочный ВОМ.')
    ensure_no_pending(settings.registry_path)
    registry = read_registry(settings.registry_path)
    project = select_project(registry, project_name)
    passport_root, configured = settings.project_paths(project_name)
    template = find_template(passport_root, configured, project)
    template_digest = sha256(template)
    inspect_template(template, project)
    if sha256(template) != template_digest:
        raise PassportError('Первое приложение изменилось во время проверки.')
    vom = read_vom(vom_path)
    if vom.path in {registry.path, template.resolve()}:
        raise PassportError('В качестве ВОМ выбран реестр или эталонное приложение.')
    day = today or date.today()
    folder = passport_root / f'{project.next_number.replace("/", "-")} от {day:%d.%m.%Y}'
    if folder.exists():
        raise PassportError(f'Папка нового паспорта уже существует: {folder}. Перезапись не выполняется.')
    # No workbook save or passport/directory creation during validation.
    with locked_file(registry.path, writable=True) as data:
        if hashlib.sha256(data).hexdigest() != registry.digest:
            raise PassportError('Реестр изменился во время проверки.')
    with locked_file(vom.path), locked_file(template):
        pass
    writable_directory(passport_root)
    writable_directory(registry.path.parent)
    return Plan(registry, project, vom, template, template_digest, passport_root, folder, day)


def _journal(path, data):
    temporary = path.with_suffix('.tmp')
    temporary.write_text(json.dumps(data, ensure_ascii=False, indent=2, default=str), encoding='utf-8')
    if path.exists():
        atomic_replace(temporary, path)
    else:
        os.replace(temporary, path)


def generate(settings, plan, *, fail_before_commit=False):
    if plan.day != date.today():
        raise PassportError('Дата Windows изменилась после проверки. Повторите «Проверить».')
    ensure_no_pending(plan.registry.path)
    if plan.folder.exists():
        raise PassportError(f'Папка уже существует: {plan.folder}')
    token = uuid.uuid4().hex
    stage = plan.registry.path.parent / '.pasport-transactions' / token
    pending_folder = plan.passport_root / f'.pasport-pending-{token}'
    stage_registry = stage / plan.registry.path.name
    journal = stage / 'journal.json'
    backup = plan.registry.path.parent / 'Резервные копии ВЗМК' / f'{plan.registry.path.stem}_{datetime.now():%Y%m%d_%H%M%S}_{token[:8]}.xlsx'
    state = {'phase': 'prepared', 'registry': str(plan.registry.path), 'passport': str(plan.folder),
             'backup': str(backup), 'pending_folder': str(pending_folder), 'original_hash': plan.registry.digest}
    committed = False
    commit_attempted = False
    result = None
    try:
        with ExitStack() as stack:
            # The registry lock serializes simultaneous runs; recheck state after acquiring it.
            original = stack.enter_context(locked_file(plan.registry.path, writable=True))
            ensure_no_pending(plan.registry.path)
            vom_bytes = stack.enter_context(locked_file(plan.vom.path))
            template_bytes = stack.enter_context(locked_file(plan.template))
            for payload, expected, label in [(original, plan.registry.digest, 'ВЗМК.xlsx'),
                                               (vom_bytes, plan.vom.digest, 'ВОМ'),
                                               (template_bytes, plan.template_digest, 'Первое приложение')]:
                if hashlib.sha256(payload).hexdigest() != expected:
                    verb = 'изменилось' if label == 'Первое приложение' else 'изменился'
                    raise PassportError(f'{label} {verb} после проверки. Формирование остановлено.')
            if plan.folder.exists():
                raise PassportError('Папка паспорта появилась после проверки.')
            stage.mkdir(parents=True, exist_ok=False)
            _journal(journal, state)
            backup.parent.mkdir(exist_ok=True)
            with backup.open('xb') as stream:
                stream.write(original)
                stream.flush()
                os.fsync(stream.fileno())
            if sha256(backup) != plan.registry.digest:
                raise PassportError('Не совпадает контрольная сумма резервной копии.')
            stage_registry.write_bytes(original)
            pending_folder.mkdir(exist_ok=False)
            appendix = pending_folder / f'{plan.project.next_number.replace("/", "_")} - Приложение № 1.xlsx'
            appendix.write_bytes(template_bytes)
            with excel_session() as excel:
                new_row, requisites = append_row(excel, stage_registry, plan.registry, plan.project, plan.vom, plan.day)
                total_row = generate_appendix(excel, appendix, plan.project, plan.vom, requisites)
            # Read the saved outputs before publishing. No openpyxl authoring.
            staged = read_registry(stage_registry)
            found = [r for r in staged.projects[plan.project.name] if r['number'] == plan.project.next_number]
            if len(found) != 1 or found[0]['row'] != new_row or found[0]['mass'] != float(plan.vom.mass) or found[0]['connections'] != plan.vom.connections:
                raise PassportError('Проверка сохранённой строки реестра не пройдена.')
            verify_saved(stage_registry, backup, appendix, template_bytes, plan, new_row, requisites)
            state.update(phase='validated', new_row=new_row, total_row=total_row,
                         staged_registry_hash=sha256(stage_registry), appendix_hash=sha256(appendix))
            _journal(journal, state)
            if fail_before_commit:
                raise PassportError('Тестовая остановка до публикации.')
            if plan.folder.exists():
                raise PassportError('Конфликт папки перед публикацией.')
            # Persist intent first: a power loss cannot be mistaken for a safe retry.
            state['phase'] = 'commit_started'
            _journal(journal, state)
            commit_attempted = True
            atomic_replace(stage_registry, plan.registry.path, backup=stage / 'original_before_replace.xlsx')
            committed = True
            # Hold the new registry too while publishing the second filesystem object.
            replacement = stack.enter_context(locked_file(plan.registry.path, writable=True))
            if hashlib.sha256(replacement).hexdigest() != state['staged_registry_hash']:
                raise PassportError('Реестр изменился в момент публикации. Требуется ручная проверка.')
            pending_folder.rename(plan.folder)  # Windows fails if target already exists.
            state['phase'] = 'complete'
            _journal(journal, state)
            result = {'number': plan.project.next_number, 'row': new_row, 'folder': str(plan.folder),
                      'appendix': str(plan.folder / appendix.name), 'backup': str(backup), 'journal': str(journal)}
    except Exception as exc:
        if commit_attempted and not committed:
            try:
                committed = sha256(plan.registry.path) != plan.registry.digest
            except Exception:
                committed = True  # Unknown state must block retries.
        if stage.exists():
            state['phase'] = 'needs_review' if committed else 'aborted'
            state['error'] = str(exc)
            _journal(journal, state)
        if committed:
            raise PassportError(f'Реестр записан, но публикация требует проверки. Не повторяйте формирование. Резервная копия: {backup}. Журнал: {journal}. Причина: {exc}') from exc
        if isinstance(exc, PassportError):
            raise
        raise PassportError(f'Формирование остановлено до замены рабочего реестра: {exc}') from exc
    return result
