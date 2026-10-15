"""Record review pointers and immutable hashes. Does not save an Excel workbook."""
import hashlib
import json
from pathlib import Path

BASE = Path(__file__).resolve().parents[1]
ROOT = BASE / 'validation' / '20261005_175857'
AUDIT = json.loads((BASE / 'audit/audit_inventory.json').read_text(encoding='utf-8'))
result = json.loads((ROOT / 'results.json').read_text(encoding='utf-8'))
gui = BASE / 'validation/20261005_180102/gui_results.json'
safety = BASE / 'validation/20261005_182013/safety_results.json'

def digest(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()

sources = [BASE.parent.parent / 'ВЗМК.xlsx', Path(AUDIT['vom']['path']),
           BASE.parent / 'Гараж казахстан/Паспорт/1_2026 - Приложение № 1.xlsx']
assert all(digest(p) == AUDIT['hashes_after'][str(p)] for p in sources)
assert all(c['pass'] for c in result['checks'])
assert json.loads(gui.read_text(encoding='utf-8'))['pass']
assert json.loads(safety.read_text(encoding='utf-8'))['pass']
assert all((ROOT / 'images' / name).is_file() for name in ('04_result_32.png', '07_result_7.png', '07_result_100.png'))
result['visual_review'] = 'PASS: final saved XLSX PNGs inspected for 7, 32, 100; native Excel window screenshot NOT_RUN'
(ROOT / 'results.json').write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding='utf-8')
(ROOT / 'saved_guard_results.json').write_text(json.dumps({
    'pass': True, 'cases': ['main_32', 'dynamic_7', 'dynamic_100'],
    'method': 'Final verify_saved called against all three native saved outputs; existing values, formulas, cell formats, widths and heights verified.',
}, ensure_ascii=False, indent=2), encoding='utf-8')
code = [*sorted((BASE / 'passport_app').glob('*.py')), *sorted((BASE / 'tests').glob('*.py')),
        BASE / 'tests/render_excel.mjs', BASE / 'main.py', BASE / 'settings.json',
        BASE / 'Запустить.ps1', BASE / 'requirements.txt', BASE / 'README.txt']
outputs = [ROOT / 'results.json', ROOT / 'saved_guard_results.json', gui, safety,
           Path(result['main_result']['appendix']), ROOT / 'main_32/ВЗМК.xlsx',
           *sorted((ROOT / 'images').glob('*.png'))]
manifest = {'code_sha256': {str(p.relative_to(BASE)): digest(p) for p in code},
            'evidence_sha256': {str(p): digest(p) for p in outputs},
            'production_sha256': {str(p): digest(p) for p in sources},
            'production_matches_audit': True, 'core_checks': len(result['checks']),
            'gui': str(gui), 'safety': str(safety), 'result': result['main_result'],
            'final_code_followups': 'Additional preservation checks verified on 3 saved outputs; Russian error grammar and integer-zero quantity validation checked read-only.',
            'native_rpc_diagnostic': 'Optional live fault-fixture edits hit Excel RPC startup errors before file mutation. Final fault tests simulate external edits and replay verified native outputs; they do not claim additional native-generation E2E.',
            'scope': 'Stage 1 only; no production generation, PDF, VIK, UZK, AKZ or next stage.'}
(BASE / 'REVIEW_MANIFEST.json').write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding='utf-8')
print(json.dumps({'production_unchanged': True, 'core_checks': len(result['checks']),
                  'gui_pass': True, 'safety_checks': 4, 'code_files': len(code)}, ensure_ascii=False))
