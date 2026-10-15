"""S5-R1 current-code evidence; no real or fixture Word generation."""
import ast
from datetime import datetime, timezone
import difflib
import json
from pathlib import Path
import sys

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT))
import uzk_stage5.core as C
import uzk_stage5.inputs as I
from profiles_stage4.offline import no_network

RUN = Path(__file__).resolve().parent
REVIEW = ROOT/'validation/stage5_independent_review_20261008_01'


def read(path):
    return json.loads(Path(path).read_text(encoding='utf-8-sig'))


def write(name, data):
    (RUN/name).write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding='utf8')


baseline = read(RUN/'baseline.json')
tests = read(RUN/'targeted_tests.json')
assert tests['successful'] and tests['tests_run'] == 17 and not tests['skipped']
before = read(RUN/'product_before_final_tests.json')
after = read(RUN/'product_after_final_tests.json')
current = {path: I.sha(path) for path in before}
assert before == after == current
I.verify_pins()

# Required regression is READ-ONLY preview, never generate.
implementation = read(ROOT/'validation/stage5_implementation_20261008_105100/EVIDENCE_MANIFEST.json')
old_run = read(ROOT/'validation/stage5_implementation_20261008_105100/final_run.json')
with no_network() as attempts:
    plan = C.preview('Гараж Казахстан', '9/2026', old_run['preview']['vom']['path'])
plan.verify()
assert not attempts
expected = {'Бп2-1': (8, [('1280', 5), ('1024', 8.4)]),
            'Пс1-9': (9, [('3600', 4)]), 'Пс1-10': (15, [('6000', 4)]),
            'Пс2-16': (1, [('480', 3)])}
assert plan.data['status'] == 'UZK_REQUIRED'
assert len(plan.data['marks']) == 4 and plan.data['controlled_units'] == 33
assert plan.data['vom']['visible_count'] == 104 and plan.data['vom']['mass_kg'] == '15822.14'
for mark in plan.data['marks']:
    q, segments = expected[mark['mark']]
    assert mark['Q'] == q and [(s['L'], s['S']) for s in mark['segments']] == segments
    assert all(s['frequency_mhz'] == 5 and s['angle_degrees'] == 70 for s in mark['segments'])
write('actual_four_mark_preview.json', {'data': plan.data, 'seal': plan.seal,
    'kind': 'CURRENT_CODE_READ_ONLY_PREVIEW', 'generation': 'NOT_RUN',
    'network_attempts': attempts})

docx = baseline['existing_docx']
assert docx == implementation['final_output']
assert I.sha(docx['path']) == docx['sha256'] == '7723be279f6e776a472fb33f0b1bc153aecd14fc04e5b319fae73f89f2960661'
write('existing_docx_integrity.json', {'path': docx['path'], 'before_sha256': docx['sha256'],
    'after_sha256': I.sha(docx['path']), 'unchanged': True, 'regenerated': False,
    'native_word_open': 'NOT_RUN; unchanged pre-fix artifact, not current-code generation proof'})

changed_runtime = {str(ROOT/'uzk_stage5'/name) for name in ('core.py', 'inputs.py', 'gui.py')}
protected_after = {path: I.sha(path) for path in baseline['protected_before']}
changes = [path for path, h in protected_after.items() if h != baseline['protected_before'][path]]
assert set(changes) == changed_runtime
write('source_after.json', protected_after)
write('source_preservation.json', {'expected_product_changes': sorted(changed_runtime),
    'unexpected_changes': [], 'unchanged_count': len(protected_after) - len(changes),
    'previous_work_manifest_unchanged': I.sha(REVIEW/'REVIEW_MANIFEST.json') == baseline['review_manifest_sha256'],
    'all_previous_work_evidence_unchanged': True, 'source_data_mutations': [],
    'production_registry_access': [], 'new_docx': 0})

def functions(path):
    tree = ast.parse(Path(path).read_text(encoding='utf-8-sig'))
    return {node.name: ast.dump(node, include_attributes=False) for node in tree.body
            if isinstance(node, (ast.FunctionDef, ast.ClassDef))}

old_core = REVIEW/'isolated_base/uzk_stage5/core.py'
old_inputs = REVIEW/'isolated_base/uzk_stage5/inputs.py'
assert I.sha(old_core) == baseline['protected_before'][str(ROOT/'uzk_stage5/core.py')]
assert I.sha(old_inputs) == baseline['protected_before'][str(ROOT/'uzk_stage5/inputs.py')]
ca, cb = functions(old_core), functions(ROOT/'uzk_stage5/core.py')
ia, ib = functions(old_inputs), functions(ROOT/'uzk_stage5/inputs.py')
core_changed = [name for name in ca if ca[name] != cb[name]]
inputs_changed = [name for name in ia if ia[name] != ib[name]]
assert core_changed == ['preview'] and inputs_changed == ['mark_profiles']
assert I.sha(ROOT/'uzk_stage5/word.py') == baseline['protected_before'][str(ROOT/'uzk_stage5/word.py')]
write('scope_check.json', {'core_functions_changed': core_changed, 'inputs_functions_changed': inputs_changed,
    'gui_change': 'only no-UZK status messages', 'calculate_marks_unchanged': True,
    'generate_unchanged': True, 'word_renderer_unchanged': True, 'Stage_1_4_unchanged': True})
patches = []
for name in ('core.py', 'inputs.py', 'gui.py'):
    old = REVIEW/'isolated_base/uzk_stage5'/name
    new = ROOT/'uzk_stage5'/name
    patches.extend(difflib.unified_diff(old.read_text(encoding='utf8').splitlines(True),
        new.read_text(encoding='utf8').splitlines(True), fromfile='BEFORE/uzk_stage5/'+name,
        tofile='AFTER/uzk_stage5/'+name))
(RUN/'S5-R1.patch').write_text(''.join(patches), encoding='utf8')
write('product_current.json', current)
assert current == {path: I.sha(path) for path in before}

not_run = ['Real DEV Word regeneration', 'Native Word revalidation of unchanged old DOCX',
    'User GUI E2E; programmatic no-UZK UI gate tested', 'Full Stage 1–4 regression',
    'Production registry reads/hashes/writes', 'Registry decrement/delete/close',
    'Web / PDF / EXE / deploy / Stage 6']
report = f'''1. ACTUAL STATE
Reviewed baseline matched all 75 source/product fingerprints, 519 prior Work evidence
files, the original implementation manifest and exact existing TEST DOCX BEFORE edits.
S5-R1 reproduced on a new copy of retained fixture bytes: 25 valid registry rows,
zero selected PROJECT+MARK matches, MARK_PROFILE_CONFLICT at unrelated row 198.

2. ROOT CAUSE S5-R1
preview loaded/validated all project profile mappings before establishing the actual
current-shipment UZK intersection. Existing no-UZK test bypassed full preview.

3. FIX IMPLEMENTATION
After valid VOM + mass and mandatory valid registry read, preview computes the normalized
PROJECT+MARK intersection. Empty intersection skips mark_profiles and Stage 4 loading,
then returns a sealed status NO_UZK_REQUIRED with zero marks / zero controlled units.
Nonempty intersection passes only required marks to mark_profiles. Its no-filter default
remains strict. Required mappings and all existing registry/VOM/checkpoint validation
remain fail-closed. Geometry, quantity arithmetic and generate are unchanged.
UI message: Для текущего паспорта УЗК не требуется. Generation button stays disabled.

4. CHANGED FILES
Modified only uzk_stage5/core.py (preview), inputs.py (mark_profiles filter),
gui.py (two no-UZK messages). Added tests/test_stage5_s5_r1.py and this validation run.
Word renderer/layout, Stage 1–4, formulas, sources and prior evidence unchanged.

5. NO-UZK PREVIEW RESULT
NO_UZK_REQUIRED; marks=[]; controlled_units=0; sealed preview.verify PASS.
Fixture A includes two unrelated conflicting identities plus ПРОВЕРИТЬ / missing profile.
Spies prove mark_profiles and load_catalog not called for zero intersection.
C: selected project registry marks absent from VOM -> same status.
D: same mark belongs only to another project -> same status.

6. NO-UZK PREVIEW -> GENERATE TEST
Full B pipeline: explicit output directory and default output both return None.
New DOCX: 0. New output directories: 0. Renderer calls: 0.
Complete fixture file hashes and directory set unchanged after preview/generate.

7. NEGATIVE REGRESSION
PASS: missing registry, unreadable registry, invalid header, invalid quantity,
required registry identity conflict, invalid VOM, required profile mapping conflict,
missing/ПРОВЕРИТЬ/NOT_FOUND profile, actual Stage 4 AMBIGUOUS/CONFLICT and unsupported
rule all STOP. Valid required marks still process with unrelated conflicting rows.
Mandatory DEV and registry checks were not weakened or bypassed by the product.

8. FOUR-MARK REGRESSION
Fresh current-code READ-ONLY preview: Гараж Казахстан / 9/2026; 104 visible rows,
15822.14 kg. Бп2-1 Q=8 ->1280×5;1024×8.4. Пс1-9 Q=9 ->3600×4.
Пс1-10 Q=15 ->6000×4. Пс2-16 Q=1 ->480×3. Four marks / 33 units;
every segment 5 MHz /70 degrees. No cross-mark merge. No Word generation.

9. EXISTING TEST DOCX INTEGRITY
Exact pre-fix artifact: {docx['path']}
SHA256 BEFORE=AFTER: {docx['sha256']}
Not recreated, not opened for writing, not represented as output of the fixed code.

10. TARGETED VALIDATION
17/17 fresh isolated tests PASS; zero skips. Tests executed against current runtime
product, not the old fixture copies. Cases A–H and required error boundaries covered.
First harness attempt was stopped by the write guard because openpyxl requested a
system TEMP file; no outside write occurred. TEMP was redirected inside this run,
the failed attempt retained, and fresh fixtures passed. Not reused as PASS evidence.
Python network attempts: 0; production reads and writes outside run: 0 in access audit.

11. EVIDENCE INTEGRITY
Product BEFORE final tests = AFTER = CURRENT, including the new targeted test script.
All data inputs / accepted stages / old documents and all previous Work evidence
unchanged. Scope check proves calculate_marks/generate AST and Word renderer unchanged.
New REVIEW_MANIFEST.json written LAST, binds every evidence file with SHA/size;
does not hash itself. Fresh preview and this test result prove current fixed bytes.

12. REGRESSION IMPACT
Only profile selection/validation scope and directly related empty-preview UI messages.
No formula, frequency/angle, layout, registry decrement/close or Stage 1–4 changes.

13. NOT_RUN
{chr(10).join('- '+item for item in not_run)}

14. READY FOR TARGETED INDEPENDENT REREVIEW — YES
S5-R1 fix completed. Work stopped; Stage 6 not started.
'''
(RUN/'REPORT.txt').write_text(report, encoding='utf8')
evidence = {str(p.relative_to(RUN)): {'sha256': I.sha(p), 'bytes': p.stat().st_size}
            for p in sorted(RUN.rglob('*')) if p.is_file() and p.name != 'REVIEW_MANIFEST.json'}
manifest = {'created_utc': datetime.now(timezone.utc).isoformat(), 'finding': 'S5-R1',
    'status': 'READY_FOR_TARGETED_INDEPENDENT_REREVIEW', 'reviewed_baseline': baseline,
    'product_before_final_tests': before, 'product_after': after, 'product_current': current,
    'tests': tests, 'actual_four_mark_regression': 'READ_ONLY_PREVIEW_PASS',
    'existing_test_docx': docx, 'existing_test_docx_unchanged': True,
    'runtime_changes': sorted(changed_runtime), 'unexpected_source_changes': [],
    'previous_work_evidence_unchanged': True, 'new_docx': 0, 'new_output_directories': 0,
    'not_run': not_run, 'evidence_files': evidence, 'manifest_written_last': True}
write('REVIEW_MANIFEST.json', manifest)  # LAST FILE WRITE
print(json.dumps({'status': manifest['status'], 'tests_passed': tests['tests_run'],
    'existing_docx_sha256': docx['sha256'], 'evidence_files': len(evidence),
    'manifest_sha256': I.sha(RUN/'REVIEW_MANIFEST.json')}, ensure_ascii=False))
