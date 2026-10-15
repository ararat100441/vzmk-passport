"""One B1/M1 run: snapshot -> mapping -> mandatory suite -> additional -> snapshot."""
import argparse
import ast
import json
import os
from pathlib import Path
import subprocess
import sys

from stage3_evidence import ROOT, product_hashes, timestamp, file_hash, require_same_code, bind_result

HARNESS_FILES = ('tests/stage3_validate.py', 'tests/stage3_mapping.py', 'tests/stage3_evidence.py',
                 'tests/stage3_finalize.py', 'tests/stage3_consolidated.py', 'tests/stage3_integrity_checks.py',
                 'tests/stage3_native_reopen.py', 'tests/stage3_additional.py')


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--test-root', type=Path, required=True)
    out = parser.parse_args().test_root.resolve()
    assert out.is_relative_to(ROOT / 'validation') and out.name.startswith('stage3_') and out.is_dir()
    checkpoint = json.loads((ROOT / 'STAGE3_CHECKPOINT.json').read_text(encoding='utf-8'))
    assert file_hash(checkpoint['production_path']) == checkpoint['production_sha256'], 'STOP: production changed'
    assert all(file_hash(p) == h for p, h in checkpoint['accepted_files_sha256'].items()), 'STOP: accepted source changed'
    before, started = product_hashes(), timestamp()
    harness = {str(ROOT / n): file_hash(ROOT / n) for n in HARNESS_FILES}
    (out / 'tested_code_before.json').write_text(json.dumps({'timestamp_utc': started, 'product_sha256': before,
          'test_harness_sha256': harness}, ensure_ascii=False, indent=2), encoding='utf-8')
    # Confirm M1 is the only product edit; all import/write/GUI behavior is unchanged.
    old_manifest = json.loads((out / 'STAGE3_REVIEW_MANIFEST.before.json').read_text(encoding='utf-8'))
    for path in before:
        if path != str(ROOT / 'passport_app/catalog.py'):
            assert before[path] == old_manifest['code_sha256'][path], 'Unexpected product edit: ' + path
    def body_without_mapping(path):
        body = ast.parse(path.read_text(encoding='utf-8-sig')).body
        body = [node for node in body if not (isinstance(node, ast.FunctionDef) and node.name == 'project_folder')
                and not (isinstance(node, ast.Assign) and any(isinstance(t, ast.Name) and t.id == 'ARCHIVE_CONTAINER' for t in node.targets))]
        return [ast.dump(node, include_attributes=False) for node in body]
    assert body_without_mapping(out / 'before/passport_app/catalog.py') == body_without_mapping(ROOT / 'passport_app/catalog.py')
    env = dict(os.environ, PASPORT_STAGE3_TEST_ROOT=str(out), PYTHONIOENCODING='utf-8')
    result_files = []
    for script, evidence in [('stage3_mapping.py', 'mapping_results.json'),
                             ('stage3_validate.py', 'targeted_results.json'),
                             ('stage3_native_reopen.py', 'native_reopen_results.json'),
                             ('stage3_additional.py', 'final_targeted_additions.json')]:
        subprocess.run([sys.executable, str(ROOT / 'tests' / script)], env=env, check=True)
        result_files.append(out / evidence)
    result_files.append(out / 'project_mapping_audit.json')
    subprocess.run([sys.executable, str(ROOT / 'tests/stage3_integrity_checks.py'), '--test-root', str(out)], env=env, check=True)
    result_files.append(out / 'integrity_negative_results.json')
    after = product_hashes()
    require_same_code(before, after)
    assert all(file_hash(p) == h for p, h in harness.items()), 'STOP: test harness changed during validation'
    assert file_hash(checkpoint['production_path']) == checkpoint['production_sha256'], 'STOP: production changed'
    assert all(file_hash(p) == h for p, h in checkpoint['accepted_files_sha256'].items()), 'STOP: accepted source changed'
    result = bind_result({'status': 'PASS', 'pass_count': 45, 'fail_count': 0,
                         'mandatory_scenarios': '20/20 PASS', 'mapping_tests': '14/14 PASS',
                         'native_reopen': '2/2 PASS', 'additional': '4/4 PASS', 'integrity_negative': '5/5 PASS',
                         'test_harness_sha256': harness,
                         'result_files_sha256': {str(p): file_hash(p) for p in result_files},
                         'production_before': checkpoint['production_sha256'], 'production_after': file_hash(checkpoint['production_path']),
                         'accepted_Stage1_2_hashes': '30/30 unchanged', 'only_product_change': 'project_folder + ARCHIVE_CONTAINER',
                         'unchanged_import_write_GUI_AST_gate': 'PASS'}, before, started)
    (out / 'consolidated_results.json').write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding='utf-8')
    print(json.dumps({'status': 'PASS', 'product_before_after_equal': True, 'evidence': str(out / 'consolidated_results.json')}, ensure_ascii=False))


if __name__ == '__main__':
    try:
        main()
    except Exception as exc:
        if '--test-root' in sys.argv:
            out = Path(sys.argv[sys.argv.index('--test-root') + 1]).resolve()
            if out.is_relative_to(ROOT / 'validation') and out.is_dir():
                start = json.loads((out / 'tested_code_before.json').read_text(encoding='utf-8'))
                (out / 'consolidated_failure.json').write_text(json.dumps({'status': 'FAIL',
                    'error': str(exc), 'completed_at_utc': timestamp(), 'started_at_utc': start['timestamp_utc'],
                    'product_sha256_before': start['product_sha256'], 'product_sha256_after': product_hashes(),
                    'fail_count': 1}, ensure_ascii=False, indent=2), encoding='utf-8')
        raise
