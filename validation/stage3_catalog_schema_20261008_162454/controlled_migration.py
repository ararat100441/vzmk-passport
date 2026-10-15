"""One explicit authorized real migration after exact-byte isolated PASS."""
from pathlib import Path
import sys
import json
import hashlib
from datetime import datetime, timezone
from unittest.mock import patch

RUN = Path(__file__).resolve().parent
ROOT = RUN.parents[1]
sys.path.insert(0, str(ROOT))
from passport_app import catalog_migration as migration
from uzk_stage5 import core


def sha(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def save(name, value):
    (RUN / name).write_text(json.dumps(value, ensure_ascii=False, indent=2), encoding='utf8')


assert sys.argv[1:] == ['--migrate-real'], 'explicit --migrate-real required'
assert not (RUN / 'real_migration_result.json').exists(), 'already migrated; do not repeat'
baseline = json.loads((RUN / 'baseline.json').read_text(encoding='utf8'))
tests = json.loads((RUN / 'isolated_tests.json').read_text(encoding='utf8'))
assert len(tests) >= 32 and all(t['status'] == 'PASS' for t in tests)
tested = json.loads((RUN / 'product_hashes_AFTER_TESTS.json').read_text(encoding='utf8'))
assert tested == json.loads((RUN / 'product_hashes_BEFORE_TESTS.json').read_text(encoding='utf8'))
assert all(sha(p) == expected for p, expected in tested.items()), 'tested product changed'
protected = json.loads((RUN / 'protected_before.json').read_text(encoding='utf8'))
assert all(sha(p) == expected for p, expected in protected.items()), 'protected input changed'
catalog = Path(baseline['source_catalog'])
denied, reads, writes, replaces = [], set(), set(), []


def audit(event, args):
    if event != 'open' or not args or not isinstance(args[0], (str, bytes)):
        return
    p = Path(args[0].decode() if isinstance(args[0], bytes) else args[0]).absolute()
    if p.name.casefold() in ('взмк.xlsx', 'узк.xlsm'):
        denied.append(str(p))
        raise AssertionError('PRODUCTION_REGISTRY_ACCESS_FORBIDDEN')
    mode = args[1] if len(args) > 1 else None
    flags = args[2] if len(args) > 2 else 0
    write = isinstance(mode, str) and any(c in mode for c in 'wax+') or bool(flags & (1 | 2 | 64 | 512))
    if write:
        allowed = p.is_relative_to(RUN) or p.parent == catalog.parent / '.stage3-backups' or (
            p.parent == catalog.parent and p.name.startswith('.stage3-schema-') and p.name.endswith('.tmp.xlsx'))
        if not allowed or p.suffix.casefold() == '.docx':
            raise AssertionError(f'WRITE_FORBIDDEN: {p}')
        writes.add(str(p))
    elif p.suffix.casefold() in ('.xls', '.xlsx', '.xlsm'):
        reads.add(str(p))


sys.addaudithook(audit)
original_replace = migration.atomic_replace


def logged_replace(source, target):
    assert Path(target) == catalog
    assert Path(source).parent == catalog.parent and Path(source).name.startswith('.stage3-schema-')
    replaces.append({'source': str(source), 'target': str(target)})
    return original_replace(source, target)


save('real_migration_started.json', {'utc': datetime.now(timezone.utc).isoformat(),
     'source_sha256': baseline['source_sha256'], 'isolated_tests': len(tests),
     'policy': 'one controlled migration; no real rollback; no Word; no production registries'})
with patch.object(migration, 'atomic_replace', side_effect=logged_replace):
    result = migration.migrate_catalog(catalog, baseline['source_sha256'])
save('real_migration_result.json', result)
assert len(replaces) == 1
assert sha(catalog) == result['new_sha256']
assert sha(result['backup']) == baseline['source_sha256']
assert result['assessment_before']['rows'] == result['assessment_after']['rows']
save('real_native_excel_after.json', migration.native_validate(catalog))
accepted = json.loads(Path(baseline['accepted_four_mark_preview']).read_text(encoding='utf8'))['preview']
isolated = json.loads((RUN / 'stage5_isolated_preview.json').read_text(encoding='utf8'))
plan = core.preview(accepted['passport']['project'], accepted['passport']['number'], Path(isolated['vom']['path']))
assert Path(plan.data['mark_catalog']) == catalog
assert plan.data['marks'] == isolated['marks']
assert len(plan.data['marks']) == 4 and plan.data['controlled_units'] == 33
save('stage5_real_catalog_preview.json', plan.data)
protected_after = {p: sha(p) for p in protected}
assert protected_after == protected
assert all(sha(p) == expected for p, expected in tested.items())
assert not denied
save('protected_after.json', protected_after)
save('real_operation_audit.json', {'production_registry_attempts': denied,
     'python_workbook_reads': sorted(reads), 'python_writes': sorted(writes),
     'native_atomic_replacements': replaces,
     'native_excel_paths': [result['native_excel']['path'], str(catalog), isolated['vom']['path']],
     'Word_generation': 'NOT_RUN', 'real_rollback': 'NOT_RUN'})
save('real_migration_completion.json', {'status': 'PASS', 'utc': datetime.now(timezone.utc).isoformat(),
     'records': result['assessment_after']['records'], 'marks_profiles_AKZ_preserved': True,
     'native_before_atomic_replace': 'PASS', 'native_after_replace': 'PASS',
     'Stage5_real_catalog_preview': {'marks': 4, 'items': 33},
     'protected_files_unchanged': len(protected), 'production_registry_access': 0,
     'atomic_replacement_count': len(replaces)})
print(json.dumps({'status': 'PASS', 'records': result['assessment_after']['records'],
                  'new_sha256': result['new_sha256'], 'backup': result['backup'],
                  'native_excel': 'PASS', 'Stage5': '4 marks / 33 items'}, ensure_ascii=True))
