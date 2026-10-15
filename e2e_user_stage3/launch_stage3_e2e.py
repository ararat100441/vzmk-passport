"""Fixed-path launcher and CheckOnly; no replacement import implementation."""
import argparse
from datetime import datetime, timezone
import hashlib
import json
import os
from pathlib import Path
import stat
import sys

sys.stdout.reconfigure(encoding='utf-8')
sys.stderr.reconfigure(encoding='utf-8')

EXPECTED_ROOT = Path('C:/Users/arara/Desktop/ВЗМК рабочий стол/Работа/Документы/Pasport/e2e_user_stage3')
E2E = Path(__file__).resolve().parent
PROJECT = 'Гараж Казахстан'
CATALOG_NAME = 'Справочник марок.xlsx'
DEV = EXPECTED_ROOT.parent / 'DEV' / 'ВЗМК_DEV.xlsx'
CHECKPOINT = DEV.parent / 'DEV_REGISTRY_CHECKPOINT.json'
EXPECTED_SHA = '8e56f68518c70cade9efb09084dd67cd5d72e07e3fdc813b8515030e9cb98e80'
INITIAL = [('К1', '35Б1', 'RAL 7004', 0), ('К2', None, 'RAL 6027', 1), ('К4', '30Б1', 'RAL 9003', 2)]
EXPECTED_CONFIG = {'documents_root': 'Документы', 'registry_path': 'ВЗМК_DEV.xlsx',
                   'projects': {PROJECT: {'passport_root': 'Гараж казахстан/Паспорт'}}}


def sha(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def inside(path):
    path = Path(path)
    resolved = path.resolve()
    if not resolved.is_relative_to(EXPECTED_ROOT) or E2E != EXPECTED_ROOT:
        raise ValueError('STOP: path is outside fixed e2e_user_stage3: ' + str(path))
    for parent in [path] + list(path.parents):
        if parent.exists() and parent.stat(follow_symlinks=False).st_file_attributes & stat.FILE_ATTRIBUTE_REPARSE_POINT:
            raise ValueError('STOP: reparse/symlink/junction in E2E path: ' + str(parent))
        if parent == EXPECTED_ROOT:
            break
    return resolved


def protected_unchanged(baseline):
    fix = baseline.get('gui_input_preview_fix')
    gui_source = E2E.parent / 'passport_app' / 'catalog_gui.py'
    if fix:
        manifest_path = inside(fix['manifest_path'])
        if sha(manifest_path) != fix['manifest_sha256']:
            raise ValueError('STOP: targeted GUI evidence manifest changed')
        manifest = json.loads(manifest_path.read_text(encoding='utf-8'))
        if (manifest['targeted_validation'] != 'PASS'
                or manifest['product_hashes_before_tests'] != manifest['product_hashes_after_tests']
                or manifest['product_hashes_after_tests'][str(gui_source)] != fix['current_gui_sha256']
                or manifest['original_gui_sha256'] != baseline['accepted_files_sha256'][str(gui_source)]):
            raise ValueError('STOP: targeted GUI checkpoint differs')
    for path, expected in baseline['accepted_files_sha256'].items():
        source = Path(path).resolve()
        if not source.is_relative_to(E2E.parent) or source.is_relative_to(E2E) or source.is_relative_to(DEV.parent):
            raise ValueError('STOP: accepted path is outside source/evidence scope')
        current_expected = fix['current_gui_sha256'] if fix and source == gui_source else expected
        if sha(path) != current_expected:
            raise ValueError('STOP: accepted source/evidence changed: ' + path)
    if sha(CHECKPOINT) != baseline['dev_checkpoint_sha256']:
        raise ValueError('STOP: DEV checkpoint changed. Explicit user DEV update is required.')
    checkpoint = json.loads(CHECKPOINT.read_text(encoding='utf-8'))
    if (checkpoint['path'] != str(DEV) or checkpoint['sha256'] != EXPECTED_SHA
            or checkpoint['purpose'] != 'DEVELOPMENT_REGISTRY' or checkpoint['source'] != 'USER_CREATED_COPY'
            or checkpoint['auto_sync_with_production'] is not False
            or checkpoint['policy']['production_hash_gate'] is not False):
        raise ValueError('STOP: DEV checkpoint/policy differs')
    if sha(DEV) != EXPECTED_SHA:
        raise ValueError('STOP: DEV changed. Ask the user to explicitly replace/update DEV; no auto-refresh.')


def preflight():
    if not sys.flags.isolated or not sys.dont_write_bytecode:
        raise ValueError('STOP: use the E2E launcher (-I -B), with bytecode disabled')
    inside(E2E)
    if E2E != EXPECTED_ROOT:
        raise ValueError('STOP: launcher root differs from fixed E2E root')
    # Reject reparse points anywhere in the prepared environment before any GUI opens.
    for folder, directories, files in os.walk(E2E):
        for name in directories + files:
            inside(Path(folder) / name)
    baseline = json.loads(inside(E2E / 'baseline.json').read_text(encoding='utf-8'))
    if baseline['e2e_root'] != str(E2E):
        raise ValueError('STOP: baseline belongs to another E2E directory')
    if (baseline.get('development_registry_source') != str(DEV)
            or baseline.get('development_registry_sha256') != EXPECTED_SHA
            or baseline.get('dev_checkpoint_path') != str(CHECKPOINT)
            or baseline.get('production_hash_gate') is not False
            or baseline.get('auto_sync_with_production') is not False
            or any(key in baseline for key in ('production_path', 'production_sha256', 'protected_before'))):
        raise ValueError('STOP: baseline is not the prepared DEV environment')
    protected_unchanged(baseline)
    for destination, item in baseline['runtime_copies'].items():
        relative = inside(destination).relative_to(inside(E2E / 'runtime'))
        source = (E2E.parent / relative).resolve()
        if not source.is_relative_to(E2E.parent) or str(source) != item['source']:
            raise ValueError('STOP: runtime source differs')
        if sha(inside(destination)) != item['sha256'] or sha(source) != item['sha256']:
            raise ValueError('STOP: accepted runtime copy/source changed: ' + destination)
    config = inside(E2E / 'settings.json')
    if sha(config) != baseline['settings_sha256'] or json.loads(config.read_text(encoding='utf-8')) != EXPECTED_CONFIG:
        raise ValueError('STOP: E2E settings changed; production fallback is unavailable')
    paths = {'Registry path': inside(E2E / 'ВЗМК_DEV.xlsx'),
             'Documents root': inside(E2E / 'Документы'),
             'Project folder': inside(E2E / 'Документы' / 'Гараж казахстан'),
             'Catalog path': inside(E2E / 'Документы' / 'Гараж казахстан' / CATALOG_NAME),
             'Backup root': inside(E2E / 'Документы' / 'Гараж казахстан' / '.stage3-backups'),
             'Temp root': inside(E2E / 'Документы' / 'Гараж казахстан'),
             'Catalog lock path': inside(E2E / 'Документы' / 'Гараж казахстан' / '.stage3-catalog.lock'),
             'Import text path': inside(E2E / 'E2E_ИМПОРТ.txt'),
             'Runtime root': inside(E2E / 'runtime'), 'State root': inside(E2E / '.state')}
    if sha(paths['Registry path']) != EXPECTED_SHA:
        raise ValueError('STOP: E2E registry copy changed')
    if sha(paths['Catalog path']) != baseline['catalog_initial_sha256']:
        raise ValueError('STOP: E2E catalog already changed. Keep your result; no fixture is reset automatically.')
    if sha(inside(E2E / 'E2E_ИМПОРТ.txt')) != baseline['import_sha256']:
        raise ValueError('STOP: prepared import text changed')
    if sha(inside(baseline['original_fixture_path'])) != baseline['original_fixture_sha256']:
        raise ValueError('STOP: original fixture snapshot changed')
    if list(paths['Backup root'].iterdir()) or list(paths['Project folder'].glob('.stage3-*.tmp.xlsx')):
        raise ValueError('STOP: E2E already has import backup/temp state. Nothing is overwritten.')
    sys.path.insert(0, str(paths['Runtime root']))
    from passport_app.config import Settings
    from passport_app.catalog import project_folder, read_catalog, CatalogService
    from passport_app.catalog_gui import CatalogApplication
    import stage3
    for name, module in tuple(sys.modules.items()):
        if name == 'stage3' or name.startswith('passport_app'):
            if module and getattr(module, '__file__', None):
                if not Path(module.__file__).resolve().is_relative_to(paths['Runtime root']):
                    raise ValueError('STOP: product import is outside the E2E runtime: ' + name)
    settings = Settings.load(config)
    if settings.registry_path != paths['Registry path'] or settings.documents_root != paths['Documents root']:
        raise ValueError('STOP: active settings point outside the fixed E2E paths')
    if project_folder(settings, PROJECT) != paths['Project folder']:
        raise ValueError('STOP: active project mapping differs from the E2E copy')
    read_catalog(paths['Catalog path'].read_bytes())
    service = CatalogService(settings)
    text = (E2E / 'E2E_ИМПОРТ.txt').read_text(encoding='utf-8-sig')
    plan = service.preview(PROJECT, text)
    if plan.path != paths['Catalog path'] or [r.status for r in plan.rows] != ['SAME', 'FILL', 'NEW', 'CONFLICT', 'CHECK_REQUIRED']:
        raise ValueError('STOP: active E2E catalog/expected preview differs')
    if not settings.documents_root.resolve().is_relative_to(E2E):
        raise ValueError('STOP: write root outside E2E')
    print('USER E2E — ТЕСТОВАЯ КОПИЯ DEV. Все записи только внутри E2E.', flush=True)
    print('Development registry source: ' + str(DEV), flush=True)
    for label, path in paths.items():
        print(label + ': ' + str(path), flush=True)
    print('DEV checkpoint SHA-256: ' + EXPECTED_SHA, flush=True)
    print('Production: не используется; production hash gate отсутствует; auto-sync false.', flush=True)
    if baseline.get('gui_input_preview_fix'):
        print('GUI input/preview fix: targeted validation PASS; targeted Work rereview pending.', flush=True)
    return baseline, paths, settings, CatalogApplication, text


def check_only(baseline, paths, settings, Application, text):
    before = sha(paths['Catalog path'])
    inventory_before = {str(p): sha(p) for p in paths['Project folder'].rglob('*') if p.is_file()}
    app = Application(settings)
    app.withdraw()
    try:
        app.update()
        app.project_var.set(PROJECT)
        app.combo.event_generate('<<ComboboxSelected>>')
        app.input.insert('1.0', text)
        app.update()
        app.check_button.invoke()
        app.update()
        if app.plan is None or app.plan.path != paths['Catalog path']:
            raise ValueError('STOP: actual GUI preview path is not the E2E catalog')
        statuses = [r.status for r in app.plan.rows]
        if statuses != ['SAME', 'FILL', 'NEW', 'CONFLICT', 'CHECK_REQUIRED']:
            raise ValueError('STOP: actual GUI preview statuses differ')
        app.input.insert('end', '\n')
        app.update()
        if app.plan is not None or str(app.apply_button['state']) != 'disabled':
            raise ValueError('STOP: input change did not invalidate Apply')
        app.check_button.invoke()
        app.update()
        if app.plan is None:
            raise ValueError('STOP: preview recreation failed')
        alternatives = [name for name in app.combo['values'] if name != PROJECT]
        if not alternatives:
            raise ValueError('STOP: no alternate DEV project for invalidation CheckOnly')
        app.project_var.set(alternatives[0])
        app.combo.event_generate('<<ComboboxSelected>>')
        app.update()
        if app.plan is not None or str(app.apply_button['state']) != 'disabled':
            raise ValueError('STOP: project change did not invalidate Apply')
    finally:
        app.destroy()
    if sha(paths['Catalog path']) != before or inventory_before != {str(p): sha(p) for p in paths['Project folder'].rglob('*') if p.is_file()}:
        raise ValueError('STOP: preview/CheckOnly wrote to the E2E project')
    protected_unchanged(baseline)
    result = {'status': 'PASS', 'timestamp_utc': datetime.now(timezone.utc).isoformat(),
              'product_imports': 'PASS: byte-identical E2E runtime', 'active_project': PROJECT,
              'paths': {k: str(v) for k, v in paths.items()}, 'write_paths_inside_E2E': True,
              'preview_statuses': statuses, 'preview_no_write': True,
              'input_change_invalidates_plan': True, 'project_change_invalidates_plan': True,
              'backup_destination_inside_E2E': True, 'original_fixture_restorable_bytes':
                  sha(paths['Catalog path']) == sha(inside(baseline['original_fixture_path'])),
              'development_registry_source': str(DEV), 'dev_sha_before_after': EXPECTED_SHA,
              'dev_checkpoint_matches': True, 'production_hash_gate': False,
              'production_access': 'NOT_PERFORMED', 'auto_sync_with_production': False,
              'accepted_files_unchanged': len(baseline['accepted_files_sha256']) - bool(baseline.get('gui_input_preview_fix')),
              'gui_input_preview_fix': baseline.get('gui_input_preview_fix'),
              'catalog_sha_before_after': before, 'Apply': 'NOT_RUN', 'USER_E2E': 'NOT_RUN'}
    inside(E2E / '.state' / 'checkonly_result.json').write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding='utf-8')
    print('CheckOnly: PASS. Apply NOT_RUN. USER E2E NOT_RUN.', flush=True)


def main():
    parser = argparse.ArgumentParser(description='Fixed isolated Stage 3 USER E2E launcher')
    parser.add_argument('--check-only', action='store_true')
    args = parser.parse_args()
    baseline, paths, settings, Application, text = preflight()
    if args.check_only:
        check_only(baseline, paths, settings, Application, text)
    else:
        app = Application(settings)
        app.title('ВЗМК — Этап 3 USER E2E — ТЕСТОВАЯ КОПИЯ')
        app.mainloop()


if __name__ == '__main__':
    try:
        main()
    except Exception as exc:
        print('STOP: ' + str(exc), file=sys.stderr)
        raise SystemExit(1)
