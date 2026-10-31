"""Verify saved outputs before refreshing review hashes; never save a workbook."""
import argparse
import hashlib
import json
import sys
from datetime import datetime, timezone
from pathlib import Path
from types import SimpleNamespace

BASE = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(BASE))
from passport_app.checks import verify_saved
from passport_app.config import Settings
from passport_app.registry import read_registry, select_project
from passport_app.vom import read_vom

ROOT = BASE / 'validation' / '20261005_175857'
CASES = ('main_32', 'dynamic_7', 'dynamic_100')


def digest(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def write_json(path, data):
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding='utf-8')


def verify_outputs(root):
    """Reconstruct each original plan from its backup and complete journal."""
    root = Path(root).resolve()
    evidence = root / 'saved_guard_results.json'
    report = {'pass': False, 'status': 'running', 'cases': [],
              'method': 'Actual passport_app.checks.verify_saved calls; read-only saved workbook verification.',
              'checked_at': datetime.now(timezone.utc).isoformat()}
    # Invalidate any older PASS before attempting a new verification.
    write_json(evidence, report)
    for name in CASES:
        item = {'case': name, 'pass': False, 'verify_saved_called': False}
        report['cases'].append(item)
        try:
            case = root / name
            journals = list((case / '.pasport-transactions').glob('*/journal.json'))
            if len(journals) != 1:
                raise ValueError(f'{name}: expected one transaction journal')
            journal = json.loads(journals[0].read_text(encoding='utf-8'))
            if journal['phase'] != 'complete':
                raise ValueError(f'{name}: transaction is not complete')
            registry_path = Path(journal['registry']).resolve()
            backup = Path(journal['backup']).resolve()
            folder = Path(journal['passport']).resolve()
            settings = Settings.load(case / 'settings.json')
            _, template = settings.project_paths('Гараж Казахстан')
            appendices = list(folder.glob('*.xlsx'))
            if len(appendices) != 1:
                raise ValueError(f'{name}: expected one saved appendix')
            appendix = appendices[0]
            paths = {'registry': registry_path, 'appendix': appendix, 'backup': backup,
                     'vom': case / 'ВОМ отгрузочный.xlsx', 'template': template,
                     'journal': journals[0]}
            item['files'] = {key: {'path': str(path), 'sha256': digest(path)} for key, path in paths.items()}
            for key, expected in [('registry', 'staged_registry_hash'), ('appendix', 'appendix_hash'),
                                  ('backup', 'original_hash')]:
                if item['files'][key]['sha256'] != journal[expected]:
                    raise ValueError(f'{name}: {key} does not match transaction hash')
            original = read_registry(backup)
            project = select_project(original, 'Гараж Казахстан')
            incoming = read_vom(paths['vom'])
            plan = SimpleNamespace(registry=original, project=project, vom=incoming)
            saved = read_registry(registry_path)
            record = next(e for e in saved.projects[project.name] if e['row'] == journal['new_row'])
            requisites = {key: record[key] for key in ('number', 'date', 'object')}
            item['new_row'] = journal['new_row']
            item['verify_saved_called'] = True
            verified = verify_saved(registry_path, backup, appendix, template.read_bytes(),
                                    plan, journal['new_row'], requisites)
            if verified is False:
                raise ValueError('verify_saved returned False')
            if any(digest(path) != item['files'][key]['sha256'] for key, path in paths.items()):
                raise ValueError(f'{name}: verification inputs changed during checking')
            item['pass'] = True
        except Exception as exc:
            item['error'] = f'{type(exc).__name__}: {exc}'
    report['pass'] = all(item['pass'] for item in report['cases'])
    report['status'] = 'complete'
    write_json(evidence, report)
    if not report['pass']:
        raise RuntimeError('Saved workbook verification failed; manifest was not refreshed')
    return report


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--root', type=Path, default=ROOT)
    parser.add_argument('--verify-only', action='store_true')
    parser.add_argument('--targeted-results', type=Path)
    args = parser.parse_args(argv)
    if args.verify_only:
        report = verify_outputs(args.root)
        print(json.dumps({'saved_verifications': len(report['cases']), 'pass': report['pass']}))
        return
    if args.root.resolve() != ROOT.resolve():
        raise ValueError('Manifest refresh supports only the accepted validation root')
    manifest_path = BASE / 'REVIEW_MANIFEST.json'
    manifest = json.loads(manifest_path.read_text(encoding='utf-8'))
    if any(digest(Path(path)) != expected for path, expected in manifest['production_sha256'].items()):
        raise ValueError('Production source differs from review checkpoint')
    targeted = None
    if args.targeted_results:
        targeted = args.targeted_results.resolve()
        if not json.loads(targeted.read_text(encoding='utf-8'))['pass']:
            raise ValueError('Targeted checks did not pass')
    report = verify_outputs(ROOT)
    code = [*sorted((BASE / 'passport_app').glob('*.py')), *sorted((BASE / 'tests').glob('*.py')),
            BASE / 'tests/render_excel.mjs', BASE / 'main.py', BASE / 'settings.json',
            BASE / 'Запустить.ps1', BASE / 'requirements.txt', BASE / 'README.txt']
    manifest['code_sha256'] = {str(p.relative_to(BASE)): digest(p) for p in code}
    manifest['evidence_sha256'] = {path: digest(Path(path)) for path in manifest['evidence_sha256']}
    if targeted:
        manifest['evidence_sha256'][str(targeted)] = digest(targeted)
        manifest['targeted_M1_M2'] = str(targeted)
    manifest['final_code_followups'] = 'M1: exact original Connections strings. M2: actual verify_saved calls for all 3 saved outputs; targeted PASS/FAIL tests. Prior full regression not rerun.'
    write_json(manifest_path, manifest)
    print(json.dumps({'production_unchanged': True, 'saved_verifications': len(report['cases']),
                      'code_files': len(code), 'pass': report['pass']}))


if __name__ == '__main__':
    main()
