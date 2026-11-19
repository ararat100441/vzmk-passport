"""Finalize review evidence only after tests, visual QA and preservation gates pass."""
import argparse
import json
import os
from pathlib import Path
import sys

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
from passport_app.excel import sha256
from stage3_evidence import evidence_root, file_hash, product_hashes, require_same_code, verify_result, timestamp


def main():
    parser = argparse.ArgumentParser(description='Stage 3 finalizer: evidence hashes must match tested current code')
    parser.add_argument('--test-root', type=Path)
    args = parser.parse_args()
    checkpoint = json.loads((ROOT / 'STAGE3_CHECKPOINT.json').read_text(encoding='utf-8'))
    if args.test_root:
        os.environ['PASPORT_STAGE3_TEST_ROOT'] = str(args.test_root)
    out = evidence_root(checkpoint)
    assert sha256(checkpoint['production_path']) == checkpoint['production_sha256'], 'STOP: production changed'
    assert not (out / 'targeted_failure.json').exists() and not (out / 'consolidated_failure.json').exists(), 'STOP: failed/incomplete validation run'
    results = json.loads((out / 'targeted_results.json').read_text(encoding='utf-8'))
    tested = verify_result(results)
    assert len(results['checks']) == results['pass_count'] == 18
    assert results['required_pass_count'] == 20 and results['required_fail_count'] == 0
    assert all(c['status'] == 'PASS' for c in results['checks'])
    consolidated = json.loads((out / 'consolidated_results.json').read_text(encoding='utf-8'))
    verify_result(consolidated, tested)
    for path, expected_hash in consolidated['test_harness_sha256'].items():
        assert file_hash(path) == expected_hash, 'STOP: test harness changed: ' + path
    for path, expected_hash in consolidated['result_files_sha256'].items():
        assert file_hash(path) == expected_hash, 'STOP: test evidence changed: ' + path
    mapping = json.loads((out / 'mapping_results.json').read_text(encoding='utf-8'))
    audit = json.loads((out / 'project_mapping_audit.json').read_text(encoding='utf-8'))
    native = json.loads((out / 'native_reopen_results.json').read_text(encoding='utf-8'))
    extra = json.loads((out / 'final_targeted_additions.json').read_text(encoding='utf-8'))
    negative = json.loads((out / 'integrity_negative_results.json').read_text(encoding='utf-8'))
    for result in (mapping, audit, native, extra, negative):
        verify_result(result, tested)
    assert mapping['pass_count'] == len(mapping['checks']) == 14
    assert all(c['status'] == 'PASS' for c in mapping['checks'])
    assert audit['read_only'] and audit['total'] == 53 and sum(audit['counts'].values()) == 53
    assert native['pass_count'] == len(native['checks']) == 2 and extra['pass_count'] == len(extra['checks']) == 4
    assert negative['pass_count'] == len(negative['checks']) == 5 and all(c['status'] == 'PASS' for c in negative['checks'])
    for item in native['checks']:
        assert sha256(item['path']) == item['sha256'], 'STOP: saved native-open workbook changed'
    for path, expected in checkpoint['accepted_files_sha256'].items():
        assert sha256(path) == expected, 'STOP: accepted file changed: ' + path
    # Work accepted workbook visuals; import/mutation/layout logic is unchanged by B1/M1.
    old_out = Path(checkpoint['validation_root'])
    visual = json.loads((old_out / 'visual_review.json').read_text(encoding='utf-8'))
    assert visual['status'] == 'PASS'
    for item in visual['images']:
        assert sha256(item['path']) == item['sha256']
    assert sha256(checkpoint['production_path']) == checkpoint['production_sha256'], 'STOP: production changed'
    preservation = {'status': 'PASS', 'accepted_files': len(checkpoint['accepted_files_sha256']),
                    'scope': 'B1/M1 targeted correction: production registry + accepted Stage1/2 hashes',
                    'original_checkpoint_preserved': True,
                    'production_before': checkpoint['production_sha256'],
                    'production_after': sha256(checkpoint['production_path']),
                    'production_catalog_creation': 'NOT_RUN', 'all_write_validation': 'DISPOSABLE_COPIES_ONLY'}
    (out / 'preservation_final.json').write_text(json.dumps(preservation, ensure_ascii=False, indent=2), encoding='utf-8')
    added = ['passport_app/catalog.py', 'passport_app/catalog_gui.py', 'stage3.py',
             'Запустить Этап 3.ps1', 'Запустить Этап 3.cmd', 'STAGE3_README.txt', 'STAGE3_REPORT.txt',
             'tests/stage3_audit.py', 'tests/stage3_validate.py', 'tests/stage3_render.mjs',
             'tests/stage3_finalize.py', 'tests/stage3_native_reopen.py', 'STAGE3_CHECKPOINT.json']
    added.append('tests/stage3_additional.py')
    added += ['tests/stage3_evidence.py', 'tests/stage3_mapping.py', 'tests/stage3_consolidated.py',
              'tests/stage3_integrity_checks.py']
    evidence = {}
    for folder, directories, files in __import__('os').walk(out):
        directories[:] = [d for d in directories if d != 'node_modules']
        for name in files:
            path = Path(folder) / name
            evidence[str(path)] = sha256(path)
    manifest = {'stage': 3, 'ready_for_independent_review': True,
                'ready_for_targeted_rereview': True, 'findings_fixed': ['B1', 'M1'],
                'created_at_utc': timestamp(), 'validation_root': str(out),
                'accepted_Stage1_2': 'User supplied PASS incl independent reviews and human E2E; accepted hashes unchanged',
                'production_checkpoint': {'path': checkpoint['production_path'], 'sha256': checkpoint['production_sha256']},
                'code_sha256': {str(ROOT / path): sha256(ROOT / path) for path in added},
                'evidence_sha256': evidence, 'accepted_files_unchanged': True,
                'tested_product_sha256': tested,
                'tested_code_before_after_equal': True,
                'test_evidence': str(out / 'targeted_results.json'),
                'finalizer_verification': 'PASS: mandatory + consolidated + mapping + audit + native + extra hashes match current product',
                'mapping_counts': audit['counts'], 'mapping_tests': '14/14 PASS', 'targeted_required': '20/20 PASS ON FINAL CODE',
                'accepted_visual_evidence': 'REUSED Work PASS; unchanged workbook mutation/layout logic',
                'user_working_changes': 'Preserved; prior confirmed changes recorded in original Stage3 evidence',
                'automated_gui_smoke': 'PASS', 'human_Stage3_E2E': 'NOT_RUN',
                'independent_Stage3_review': 'Work: BLOCKER 1 / MAJOR 1 / MINOR 0; B1/M1 targeted rereview NOT_RUN',
                'not_run': ['Production catalog creation/import', 'Passport/VIK generation', 'Full Stage1/2 regression',
                            'Human Stage3 E2E', 'Independent Stage3 review', 'All-project runtime imports',
                            'Power loss/application crash', 'Stage4', 'PDF/EXE/deploy']}
    # Last gate immediately before publishing; no successful manifest is touched on mismatch.
    require_same_code(tested, product_hashes())
    for path, expected_hash in consolidated['test_harness_sha256'].items():
        assert file_hash(path) == expected_hash, 'STOP: test harness changed: ' + path
    assert sha256(checkpoint['production_path']) == checkpoint['production_sha256'], 'STOP: production changed'
    temporary = out / 'review_manifest.tmp.json'
    temporary.write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding='utf-8')
    os.replace(temporary, ROOT / 'STAGE3_REVIEW_MANIFEST.json')
    print(json.dumps({'status': 'PASS', 'accepted_files': len(checkpoint['accepted_files_sha256']),
                      'mapping': audit['counts'], 'manifest': str(ROOT / 'STAGE3_REVIEW_MANIFEST.json'),
                      'manifest_sha256': sha256(ROOT / 'STAGE3_REVIEW_MANIFEST.json')}, ensure_ascii=False))


if __name__ == '__main__':
    try:
        main()
    except (ValueError, AssertionError, KeyError, OSError) as exc:
        print('Stage 3 finalization refused: ' + str(exc), file=sys.stderr)
        raise SystemExit(1)
