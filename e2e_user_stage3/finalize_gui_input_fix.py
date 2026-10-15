"""Bind targeted evidence and current GUI runtime; keep historic manifests immutable."""
from datetime import datetime, timezone
import hashlib
import json
from pathlib import Path
import sys

sys.stdout.reconfigure(encoding='utf-8')
sys.stderr.reconfigure(encoding='utf-8')
E2E = Path(__file__).resolve().parent
ROOT = E2E.parent
OUT = E2E / '.state' / 'gui_input_preview_fix'
sha = lambda path: hashlib.sha256(Path(path).read_bytes()).hexdigest()


def write(path, data):
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding='utf-8')


baseline = json.loads((E2E / 'baseline.json').read_text(encoding='utf-8'))
targeted = json.loads((OUT / 'after_result.json').read_text(encoding='utf-8'))
native = json.loads((OUT / 'native_dialog_result.json').read_text(encoding='utf-8'))
gui = ROOT / 'passport_app' / 'catalog_gui.py'
assert targeted['status'] == native['status'] == 'PASS'
assert targeted['product_hashes_before'] == targeted['product_hashes_after']
assert all(sha(p) == h for p, h in targeted['product_hashes_after'].items())
assert sha(gui) == native['hashes_before'][str(gui)] == native['hashes_after'][str(gui)]
assert sha(ROOT / 'DEV' / 'ВЗМК_DEV.xlsx') == baseline['development_registry_sha256']
assert sha(OUT / 'catalog_gui.before.py') == baseline['accepted_files_sha256'][str(gui)]
for path, expected in baseline['accepted_files_sha256'].items():
    source = Path(path).resolve()
    assert source.is_relative_to(ROOT) and not source.is_relative_to(E2E)
    if source != gui:
        assert sha(source) == expected, 'Unexpected accepted source/evidence change'
for destination, item in baseline['runtime_copies'].items():
    source = Path(item['source'])
    assert source.is_relative_to(ROOT) and Path(destination).is_relative_to(E2E / 'runtime')
    assert sha(destination) == sha(source)
    if source != gui:
        assert sha(source) == item['sha256']
manifest = {'scope': 'STAGE3_GUI_INPUT_PREVIEW_ONLY', 'created_at_utc': datetime.now(timezone.utc).isoformat(),
            'product_code_changed': True, 'changed_product_files': [str(gui)],
            'original_gui_sha256': sha(OUT / 'catalog_gui.before.py'), 'current_gui_sha256': sha(gui),
            'product_hashes_before_tests': targeted['product_hashes_before'],
            'product_hashes_after_tests': targeted['product_hashes_after'],
            'targeted_validation': 'PASS', 'targeted_checks_passed': len(targeted['checks']),
            'actual_native_file_picker_and_error_dialog': 'PASS',
            'unchanged_accepted_files': len(baseline['accepted_files_sha256']) - 1,
            'historic_stage3_review_manifest_sha256': sha(ROOT / 'STAGE3_REVIEW_MANIFEST.json'),
            'development_registry_sha256': baseline['development_registry_sha256'],
            'production_access': 'NOT_PERFORMED', 'Work_targeted_rereview': 'NOT_RUN_REQUIRED',
            'Apply': 'NOT_RUN', 'USER_E2E': 'NOT_RUN',
            'reproduction_notes': 'Initial Shift+Insert false came from missing EXTENDEDKEY in test injector; corrected native test PASS. Not classified as product defect.',
            'evidence_sha256': {str(path): sha(path) for path in [
                OUT / 'catalog_gui.before.py', OUT / 'before_result.json', OUT / 'before_callback_traceback.txt',
                OUT / 'after_result.json', OUT / 'after_hashes_before.json', OUT / 'after_hashes_after.json',
                OUT / 'after_console_tracebacks.txt', OUT / 'native_dialog_result.json',
                E2E / 'gui_input_preview_validation.py', E2E / 'native_input_dialog_probe.py',
                OUT / 'catalog_gui.patch', OUT / 'REVIEW_GUI_INPUT_PREVIEW.txt']}}
review = OUT / 'GUI_INPUT_PREVIEW_REVIEW_MANIFEST.json'
write(review, manifest)
if not (OUT / 'baseline.before_gui_fix.json').exists():
    (OUT / 'baseline.before_gui_fix.json').write_bytes((E2E / 'baseline.json').read_bytes())
baseline['runtime_copies'][str(E2E / 'runtime' / 'passport_app' / 'catalog_gui.py')]['sha256'] = sha(gui)
baseline['gui_input_preview_fix'] = {'manifest_path': str(review), 'manifest_sha256': sha(review),
                                    'current_gui_sha256': sha(gui), 'targeted_rereview': 'NOT_RUN_REQUIRED'}
write(E2E / 'baseline.json', baseline)
print(json.dumps({'finalization': 'PASS', 'targeted_checks': len(targeted['checks']),
                  'current_gui_sha256': sha(gui), 'unchanged_accepted_files': manifest['unchanged_accepted_files'],
                  'targeted_rereview': 'NOT_RUN_REQUIRED'}, ensure_ascii=False))
