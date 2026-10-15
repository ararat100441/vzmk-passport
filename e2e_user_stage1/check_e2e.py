"""E2E launcher preflight only; no product changes or workbook saves."""
import hashlib
import json
import sys
from datetime import date
from pathlib import Path

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT / 'app'))
from passport_app.config import Settings
from passport_app.registry import read_registry, select_project
from passport_app.vom import read_vom


def digest(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def check():
    # Reject redirection by junctions/symlinks, including the environment root.
    if ROOT != Path(__file__).absolute().parent:
        raise ValueError('E2E root is redirected')
    checkpoint = json.loads((ROOT / 'E2E_CHECKPOINT.json').read_text(encoding='utf-8'))
    settings = Settings.load(ROOT / 'settings.e2e.json')
    fallback = Settings.load(ROOT / 'app/settings.json')
    expected_registry = ROOT / 'ВЗМК.xlsx'
    expected_docs = ROOT / 'Документы'
    expected_passports = expected_docs / 'Гараж казахстан/Паспорт'
    expected_template = expected_passports / '1_2026 - Приложение № 1.xlsx'
    for config in (settings, fallback):
        if config.registry_path != expected_registry or config.documents_root != expected_docs:
            raise ValueError('Active registry/documents paths differ from E2E-only paths')
        if set(config.projects) != {'Гараж Казахстан'}:
            raise ValueError('Unexpected E2E project mapping')
        if config.project_paths('Гараж Казахстан') != (expected_passports, expected_template):
            raise ValueError('Active template/passport paths differ from E2E-only paths')
    for path in (expected_docs, expected_passports, *[ROOT / p for p in checkpoint['copied_workbooks']],
                 *[ROOT / p for p in checkpoint['accepted_code_copies']]):
        if path.resolve() != path or not path.resolve().is_relative_to(ROOT):
            raise ValueError(f'E2E path is redirected: {path}')
    for path, expected in checkpoint['production_before'].items():
        if digest(Path(path)) != expected:
            raise ValueError(f'Production differs from accepted checkpoint: {path}')
    for relative, info in checkpoint['copied_workbooks'].items():
        if digest(ROOT / relative) != info['sha256']:
            raise ValueError('E2E workbook changed or test already performed; no automatic reset: ' + relative)
    for relative, expected in checkpoint['accepted_code_copies'].items():
        if digest(ROOT / relative) != expected:
            raise ValueError('Accepted E2E code copy changed: ' + relative)
    registry = read_registry(expected_registry)
    project = select_project(registry, 'Гараж Казахстан')
    incoming = read_vom(ROOT / 'ВОМ отгрузочный.xlsx')
    if (project.first_row, project.first_number, project.last_number, project.next_number) != (353, '1/2026', '8/2026', '9/2026'):
        raise ValueError('Unexpected E2E passport baseline')
    if (len(incoming.rows), incoming.quantity, str(incoming.mass), str(incoming.area), incoming.connections) != (32, 50, '15071.44', '477.749', 'Балка'):
        raise ValueError('Unexpected E2E VOM control values')
    result = {'pass': True, 'active_registry': str(settings.registry_path),
              'active_documents_root': str(settings.documents_root),
              'active_passport_root': str(expected_passports), 'active_template': str(expected_template),
              'test_vom': str(incoming.path), 'test_date': date.today().isoformat(),
              'expected_new_folder': str(expected_passports / f'9-2026 от {date.today():%d.%m.%Y}'),
              'first': project.first_number, 'last': project.last_number, 'next': project.next_number,
              'production_sha256': {p: digest(Path(p)) for p in checkpoint['production_before']},
              'gui_scenario': 'NOT_RUN'}
    (ROOT / 'E2E_PREFLIGHT.json').write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding='utf-8')
    print('E2E ONLY — active registry: ' + str(settings.registry_path))
    print('Expected new passport: 9/2026; date: ' + date.today().strftime('%d.%m.%Y'))
    return result


if __name__ == '__main__':
    check()
