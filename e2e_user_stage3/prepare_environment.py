"""One-time switch to user-created DEV; never invokes Apply or accesses production."""
from datetime import datetime, timezone
import hashlib
import json
from pathlib import Path
import shutil
import sys

sys.stdout.reconfigure(encoding='utf-8')
sys.stderr.reconfigure(encoding='utf-8')
E2E = Path(__file__).resolve().parent
PRODUCT = E2E.parent
DEV = PRODUCT / 'DEV' / 'ВЗМК_DEV.xlsx'
CHECKPOINT = DEV.parent / 'DEV_REGISTRY_CHECKPOINT.json'
EXPECTED_SHA = '8e56f68518c70cade9efb09084dd67cd5d72e07e3fdc813b8515030e9cb98e80'
PROJECT = 'Гараж Казахстан'
INITIAL = [('К1', '35Б1', 'RAL 7004', 0), ('К2', None, 'RAL 6027', 1), ('К4', '30Б1', 'RAL 9003', 2)]
SETTINGS = {'documents_root': 'Документы', 'registry_path': 'ВЗМК_DEV.xlsx',
            'projects': {PROJECT: {'passport_root': 'Гараж казахстан/Паспорт'}}}


def sha(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def write_json(path, data):
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2, default=str), encoding='utf-8')


def accepted_snapshot():
    files = {}
    for name in ('REVIEW_MANIFEST.json', 'STAGE2_REVIEW_MANIFEST.json', 'STAGE3_REVIEW_MANIFEST.json',
                 'STAGE3_CHECKPOINT.json', 'settings.json'):
        path = PRODUCT / name
        files[str(path)] = sha(path)
    manifest = json.loads((PRODUCT / 'STAGE3_REVIEW_MANIFEST.json').read_text(encoding='utf-8'))
    checkpoint = json.loads((PRODUCT / 'STAGE3_CHECKPOINT.json').read_text(encoding='utf-8'))
    for section in (manifest['code_sha256'], manifest['evidence_sha256'], checkpoint['accepted_files_sha256']):
        for source, expected in section.items():
            path = (PRODUCT / source).resolve()
            if not path.is_relative_to(PRODUCT) or path.is_relative_to(E2E) or path.is_relative_to(DEV.parent):
                raise ValueError('STOP: accepted path is outside source/evidence scope')
            if sha(path) != expected:
                raise ValueError('STOP: accepted file changed: ' + str(path))
            files[str(path)] = expected
    # Accepted test scripts are direct children. Do not follow the node_modules junction.
    for path in (PRODUCT / 'tests').glob('*'):
        if path.is_file():
            if not path.resolve().is_relative_to(PRODUCT / 'tests'):
                raise ValueError('STOP: test path escapes workspace')
            files[str(path)] = sha(path)
    return files


def main():
    # Use the fixed path guard without running the launcher or reading its old baseline.
    sys.path.insert(0, str(E2E))
    from launch_stage3_e2e import inside
    inside(E2E)
    if not sys.flags.isolated or not sys.dont_write_bytecode:
        raise ValueError('STOP: preparation requires Python -I -B')
    if sha(DEV) != EXPECTED_SHA:
        raise ValueError('STOP: DEV differs. Request explicit user replacement/update; no auto-refresh.')
    if CHECKPOINT.exists():
        raise ValueError('STOP: switch already prepared. Use launcher -CheckOnly; no automatic re-preparation.')
    before = accepted_snapshot()
    old = json.loads(inside(E2E / 'baseline.json').read_text(encoding='utf-8'))
    runtime = inside(E2E / 'runtime')
    for destination, item in old['runtime_copies'].items():
        relative = inside(destination).relative_to(runtime)
        source = (PRODUCT / relative).resolve()
        if not source.is_relative_to(PRODUCT) or str(source) != item['source']:
            raise ValueError('STOP: unexpected runtime source')
        if sha(source) != item['sha256'] or sha(destination) != item['sha256']:
            raise ValueError('STOP: accepted runtime differs')
    sys.path.insert(0, str(runtime))
    import openpyxl
    from passport_app.registry import read_registry, select_project
    registry = read_registry(DEV)
    if registry.sheet != 'Паспорт качества':
        raise ValueError('STOP: unexpected DEV main sheet')
    project = select_project(registry, PROJECT)
    catalog = inside(E2E / 'Документы' / 'Гараж казахстан' / 'Справочник марок.xlsx')
    book = openpyxl.load_workbook(catalog, read_only=True)
    try:
        if list(book.active.values)[1:] != INITIAL or book.active['D2'].data_type != 'n':
            raise ValueError('STOP: existing user E2E result differs. No automatic fixture reset.')
    finally:
        book.close()
    if sha(catalog) != old['catalog_initial_sha256']:
        raise ValueError('STOP: fixture bytes changed. Nothing is reset.')
    if list(inside(catalog.parent / '.stage3-backups').iterdir()) or list(catalog.parent.glob('.stage3-*.tmp.xlsx')):
        raise ValueError('STOP: E2E contains import state. Keep user results.')
    if sha(inside(old['original_fixture_path'])) != old['original_fixture_sha256']:
        raise ValueError('STOP: original fixture differs')
    timestamp = datetime.now(timezone.utc).isoformat()
    history = inside(E2E / '.history' / ('dev_switch_' + datetime.now(timezone.utc).strftime('%Y%m%d_%H%M%S')))
    history.mkdir(parents=True)
    archived = {}
    for relative in ('baseline.json', 'settings.json', 'ВЗМК.xlsx', '.state/checkonly_result.json',
                     '.state/launcher_entrypoints.json', 'КАК ПРОВЕРИТЬ ЭТАП 3.txt'):
        source = inside(E2E / relative)
        if source.exists():
            destination = inside(history / relative)
            destination.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(source, destination)
            archived[relative] = sha(source)
            if sha(destination) != archived[relative]:
                raise ValueError('STOP: history copy differs')
    write_json(history / 'ARCHIVE.json', {'active': False, 'purpose': 'HISTORICAL_E2E_ONLY_DO_NOT_EXECUTE',
                                        'files_sha256': archived})
    policy = {'development_work_validation_e2e_registry': str(DEV),
              'production_registry': str(PRODUCT.parent.parent / 'ВЗМК.xlsx'),
              'production_path_purpose': 'INFORMATION_ONLY_NEVER_OPEN_OR_HASH_FOR_DEVELOPMENT',
              'production_hash_gate': False, 'production_validation_input': False,
              'auto_sync_with_production': False,
              'dev_update_rule': 'STOP_AND_REQUEST_EXPLICIT_USER_REPLACEMENT_OR_UPDATE_NO_AUTO_REFRESH',
              'future_validation_rule': 'COPY_CHECKPOINTED_DEV_INTO_DISPOSABLE_ISOLATED_ENVIRONMENT',
              'historic_tests_and_manifests': 'IMMUTABLE_HISTORICAL_EVIDENCE_NOT_ACTIVE_DEVELOPMENT_INPUTS',
              'production_configuration': 'SEPARATE_USER_AUTHORIZED_TASK_AFTER_DEVELOPMENT'}
    checkpoint = {'path': str(DEV), 'sha256': EXPECTED_SHA, 'timestamp_utc': timestamp,
                  'purpose': 'DEVELOPMENT_REGISTRY', 'source': 'USER_CREATED_COPY',
                  'auto_sync_with_production': False, 'policy': policy,
                  'verification': {'status': 'PASS', 'sheet': registry.sheet, 'table': registry.table,
                                   'columns': registry.columns, 'project_count': len(registry.projects),
                                   'project_names': list(registry.projects),
                                   'garage_first_passport': project.__dict__}}
    copy = inside(E2E / 'ВЗМК_DEV.xlsx')
    shutil.copy2(DEV, copy)
    if sha(copy) != EXPECTED_SHA or sha(DEV) != EXPECTED_SHA:
        raise ValueError('STOP: DEV changed during copying')
    write_json(inside(E2E / 'settings.json'), SETTINGS)
    # Retire only the verified old E2E copy; its bytes are preserved in .history.
    old_copy = inside(E2E / 'ВЗМК.xlsx')
    if old_copy.exists():
        old_copy.unlink()
    write_json(CHECKPOINT, checkpoint)
    baseline = {key: old[key] for key in ('accepted_review', 'runtime_copies', 'import_sha256',
                                         'catalog_initial_sha256', 'original_fixture_path',
                                         'original_fixture_sha256', 'initial_rows')}
    baseline.update({'prepared_at_utc': timestamp, 'e2e_root': str(E2E),
                     'development_registry_source': str(DEV), 'development_registry_sha256': EXPECTED_SHA,
                     'dev_checkpoint_path': str(CHECKPOINT), 'dev_checkpoint_sha256': sha(CHECKPOINT),
                     'accepted_files_sha256': before, 'settings': SETTINGS,
                     'settings_sha256': sha(E2E / 'settings.json'), 'apply': 'NOT_RUN',
                     'runtime_state_policy': 'E2E only; Python isolated mode and bytecode disabled',
                     'production_hash_gate': False, 'auto_sync_with_production': False,
                     'history_path': str(history)})
    if accepted_snapshot() != before or sha(DEV) != EXPECTED_SHA:
        raise ValueError('STOP: accepted files or DEV changed during preparation')
    write_json(inside(E2E / 'baseline.json'), baseline)
    write_json(inside(E2E / '.state' / 'dev_switch_result.json'),
               {'status': 'PASS', 'timestamp_utc': timestamp, 'source': str(DEV), 'sha256': EXPECTED_SHA,
                'accepted_files_unchanged': len(before), 'runtime_files_unchanged': len(old['runtime_copies']),
                'fixture_preserved': True, 'history_path': str(history),
                'production_access': 'NOT_PERFORMED', 'Apply': 'NOT_RUN'})
    print('DEV switch PASS; fixture preserved; accepted files unchanged:', len(before))


if __name__ == '__main__':
    try:
        main()
    except Exception as exc:
        print(str(exc), file=sys.stderr)
        raise SystemExit(1)
