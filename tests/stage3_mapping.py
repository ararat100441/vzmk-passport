"""M1 mapping scenarios and read-only audit using the actual product resolver."""
import json
from pathlib import Path
import sys

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
from stage3_evidence import product_hashes, timestamp, evidence_root, bind_result
PRODUCT_BEFORE, STARTED_AT = product_hashes(), timestamp()
from passport_app.catalog import project_folder, ARCHIVE_CONTAINER, CATALOG
from passport_app.config import Settings, PassportError
from passport_app.registry import read_registry
from passport_app.excel import sha256

CHECKPOINT = json.loads((ROOT / 'STAGE3_CHECKPOINT.json').read_text(encoding='utf-8'))
OUT = evidence_root(CHECKPOINT)


def setup(name, paths, explicit=None):
    documents = OUT / 'mapping' / name / 'Документы'
    documents.mkdir(parents=True, exist_ok=False)
    for value in paths:
        (documents / value).mkdir(parents=True)
    return Settings(documents, OUT / 'ВЗМК.xlsx', explicit or {})


def main():
    assert sha256(CHECKPOINT['production_path']) == CHECKPOINT['production_sha256'], 'STOP: production changed'
    checks = []
    for label, project, relative in [('A', 'Гараж Казахстан', 'Гараж казахстан'),
                                     ('B', 'Гранит', 'Гранит'),
                                     ('C', 'Быков', ARCHIVE_CONTAINER + '/Быков'),
                                     ('D', 'Роксемаш', ARCHIVE_CONTAINER + '/Роксемаш'),
                                     ('E', 'Теплица', ARCHIVE_CONTAINER + '/Теплица')]:
        settings = setup(label, [relative])
        selected = project_folder(settings, project)
        assert selected == (settings.documents_root / relative).resolve()
        checks.append({'scenario': label, 'status': 'PASS', 'project': project, 'selected': str(selected)})
    rejected = [('F', 'Проект X', ['Проект X', ARCHIVE_CONTAINER + '/проект  X'], 'Неоднозначная'),
                ('G', 'ТЭМПО 1', ['ТЭМПО'], 'explicit mapping'),
                ('H', 'ТЭМПО 2', ['ТЭМПО'], 'explicit mapping'),
                ('I', 'Ковчег', ['КОВЧЕГ', ARCHIVE_CONTAINER + '/ Ковчег '], 'Неоднозначная'),
                ('L', 'café', ['café', ARCHIVE_CONTAINER + '/cafe\u0301'], 'Неоднозначная'),
                ('M_no_recursion', 'Быков', [ARCHIVE_CONTAINER + '/Группа/Быков'], 'explicit mapping'),
                ('N_no_fuzzy_Kovcheg', 'Ковчег', ['КОВЧЕГ 20 тн', ARCHIVE_CONTAINER + '/КОВЧЕГ (Пестравка) КМД1'], 'explicit mapping')]
    for label, project, paths, expected in rejected:
        settings = setup(label, paths)
        try:
            project_folder(settings, project)
        except PassportError as exc:
            assert expected in str(exc), str(exc)
            checks.append({'scenario': label, 'status': 'PASS', 'project': project, 'safe_error': str(exc)})
        else:
            raise AssertionError(f'{label}: expected STOP')
    settings = setup('J', ['Проект', ARCHIVE_CONTAINER + '/Проект', ARCHIVE_CONTAINER + '/Настроенная папка/Паспорт'],
                     {'Проект': {'passport_root': ARCHIVE_CONTAINER + '/Настроенная папка/Паспорт'}})
    selected = project_folder(settings, 'Проект')
    assert selected == (settings.documents_root / ARCHIVE_CONTAINER / 'Настроенная папка').resolve()
    checks.append({'scenario': 'J_explicit_priority', 'status': 'PASS', 'selected': str(selected)})
    settings = setup('K', ['Проект'], {'Проект': {'passport_root': 'Нет такой папки/Паспорт'}})
    try:
        project_folder(settings, 'Проект')
    except PassportError as exc:
        assert 'Не найдена папка' in str(exc)
        checks.append({'scenario': 'K_missing_explicit_no_fallback', 'status': 'PASS', 'safe_error': str(exc)})
    else:
        raise AssertionError('Missing explicit path must not fall back to the direct candidate')

    result = bind_result({'status': 'PASS', 'checks': checks, 'pass_count': len(checks), 'fail_count': 0}, PRODUCT_BEFORE, STARTED_AT)
    (OUT / 'mapping_results.json').write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding='utf-8')

    settings = Settings.load(ROOT / 'settings.json')
    registry = read_registry(settings.registry_path)
    assert registry.digest == CHECKPOINT['production_sha256'] and len(registry.projects) == 53
    inventory = []
    counts = dict.fromkeys(('AUTO_DIRECT', 'AUTO_ARCHIVE', 'EXPLICIT', 'AMBIGUOUS', 'MISSING'), 0)
    for project in registry.projects:
        try:
            folder = project_folder(settings, project)
            status = ('EXPLICIT' if project in settings.projects else 'AUTO_DIRECT' if
                      folder.parent == settings.documents_root.resolve() else 'AUTO_ARCHIVE')
            item = {'project': project, 'status': status, 'folder': str(folder),
                    'catalog_exists': (folder / CATALOG).exists()}
        except PassportError as exc:
            status = 'AMBIGUOUS' if 'Неоднозначная' in str(exc) else 'MISSING'
            item = {'project': project, 'status': status, 'reason': str(exc)}
        inventory.append(item)
        counts[status] += 1
    for project in ('Быков', 'Роксемаш', 'Теплица', 'СКФ', 'Щёкиназот'):
        assert next(i for i in inventory if i['project'] == project)['status'] == 'AUTO_ARCHIVE'
    for project in ('ТЭМПО 1', 'ТЭМПО 2', 'Ковчег'):
        assert next(i for i in inventory if i['project'] == project)['status'] in ('MISSING', 'AMBIGUOUS')
    assert sha256(settings.registry_path) == CHECKPOINT['production_sha256'], 'STOP: production changed'
    audit = bind_result({'status': 'PASS', 'read_only': True, 'total': len(inventory), 'counts': counts,
                         'projects': inventory, 'pass_count': len(inventory), 'fail_count': 0,
                         'production_before': registry.digest, 'production_after': sha256(settings.registry_path)}, PRODUCT_BEFORE, STARTED_AT)
    (OUT / 'project_mapping_audit.json').write_text(json.dumps(audit, ensure_ascii=False, indent=2), encoding='utf-8')
    print(json.dumps({'mapping_tests': f'{len(checks)}/{len(checks)} PASS', 'audit': counts, 'total': len(inventory)}, ensure_ascii=False))


if __name__ == '__main__':
    main()
