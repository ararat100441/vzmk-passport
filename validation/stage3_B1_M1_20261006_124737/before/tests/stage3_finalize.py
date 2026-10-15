"""Finalize review evidence only after tests, visual QA and preservation gates pass."""
import json
from pathlib import Path
import sys

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
from passport_app.excel import sha256
from stage3_audit import protected


def main():
    checkpoint = json.loads((ROOT / 'STAGE3_CHECKPOINT.json').read_text(encoding='utf-8'))
    out = Path(checkpoint['validation_root'])
    assert sha256(checkpoint['production_path']) == checkpoint['production_sha256'], 'STOP: production changed'
    results = json.loads((out / 'targeted_results.json').read_text(encoding='utf-8'))
    assert results['status'] == 'PASS' and len(results['checks']) == 18
    for path, expected in checkpoint['accepted_files_sha256'].items():
        assert sha256(path) == expected, 'STOP: accepted file changed: ' + path
    current = protected()
    expected = dict(checkpoint['protected_sha256'])
    acknowledged = []
    acknowledgment_path = out / 'user_working_changes_acknowledgment.json'
    if acknowledgment_path.exists():
        acknowledgment = json.loads(acknowledgment_path.read_text(encoding='utf-8'))
        assert acknowledgment['source'] == 'DIRECT_USER_CONFIRMATION'
        acknowledged = acknowledgment['changes']
        for item in acknowledged:
            assert expected.get(item['path']) == item['before']
            expected[item['path']] = item['after']
    assert current == expected, 'STOP: unacknowledged protected production documents changed'
    original_unchanged = sum(current.get(p) == h for p, h in checkpoint['protected_sha256'].items())
    visual = json.loads((out / 'visual_review.json').read_text(encoding='utf-8'))
    assert visual['status'] == 'PASS'
    for item in visual['images']:
        assert sha256(item['path']) == item['sha256']
    native = json.loads((out / 'native_reopen_results.json').read_text(encoding='utf-8'))
    assert len(native) == 2 and all(item['status'] == 'PASS' for item in native)
    for item in native:
        assert sha256(item['path']) == item['sha256']
    extra = json.loads((out / 'final_targeted_additions.json').read_text(encoding='utf-8'))
    assert extra['status'] == 'PASS'
    assert sha256(checkpoint['production_path']) == checkpoint['production_sha256'], 'STOP: production changed'
    preservation = {'status': 'PASS', 'accepted_files': len(checkpoint['accepted_files_sha256']),
                    'protected_files': len(current), 'protected_sha256': current,
                    'original_protected_files_unchanged': original_unchanged,
                    'user_confirmed_working_changes': acknowledged,
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
    evidence = {}
    for folder, directories, files in __import__('os').walk(out):
        directories[:] = [d for d in directories if d != 'node_modules']
        for name in files:
            path = Path(folder) / name
            evidence[str(path)] = sha256(path)
    manifest = {'stage': 3, 'ready_for_independent_review': True,
                'accepted_Stage1_2': 'User supplied PASS incl independent reviews and human E2E; accepted hashes unchanged',
                'production_checkpoint': {'path': checkpoint['production_path'], 'sha256': checkpoint['production_sha256']},
                'code_sha256': {str(ROOT / path): sha256(ROOT / path) for path in added},
                'evidence_sha256': evidence, 'accepted_files_unchanged': True,
                'production_protected_files_unchanged': original_unchanged,
                'user_confirmed_working_changes': acknowledged,
                'current_protected_file_count': len(current), 'targeted_required': '20/20 PASS',
                'automated_gui_smoke': 'PASS', 'human_Stage3_E2E': 'NOT_RUN',
                'independent_Stage3_review': 'NOT_RUN',
                'not_run': ['Production catalog creation/import', 'Passport/VIK generation', 'Full Stage1/2 regression',
                            'Human Stage3 E2E', 'Independent Stage3 review', 'All-project runtime imports',
                            'Power loss/application crash', 'Stage4', 'PDF/EXE/deploy']}
    (ROOT / 'STAGE3_REVIEW_MANIFEST.json').write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding='utf-8')
    print(json.dumps({'status': 'PASS', 'accepted_files': len(checkpoint['accepted_files_sha256']),
                      'original_protected_unchanged': original_unchanged,
                      'user_confirmed_working_changes': len(acknowledged),
                      'protected_files': len(current), 'manifest': str(ROOT / 'STAGE3_REVIEW_MANIFEST.json'),
                      'manifest_sha256': sha256(ROOT / 'STAGE3_REVIEW_MANIFEST.json')}, ensure_ascii=False))


if __name__ == '__main__':
    main()
