"""Finalize read-only S5-R1 rereview; all writes are inside this evidence folder."""
import ast
import hashlib
import json
from datetime import datetime, timezone
from pathlib import Path

OUT = Path(__file__).resolve().parent
ROOT = OUT.parent.parent
FIX = ROOT / 'validation/stage5_S5_R1_fix_20261008_115014'
PREVIOUS = ROOT / 'validation/stage5_independent_review_20261008_01'
MANIFEST = OUT / 'REVIEW_MANIFEST.json'
FORBIDDEN = {'взмк.xlsx', 'узк.xlsm'}


def sha(path):
    path = Path(path)
    assert path.name.casefold() not in FORBIDDEN, f'Production access forbidden: {path}'
    return hashlib.sha256(path.read_bytes()).hexdigest()


def read(path):
    return json.loads(Path(path).read_text(encoding='utf-8-sig'))


def write(path, data):
    path = Path(path).resolve()
    assert path.is_relative_to(OUT)
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2) + '\n', encoding='utf-8')


def verify_evidence(folder):
    manifest = folder / 'REVIEW_MANIFEST.json'
    data = read(manifest)
    # Prior Windows manifests use backslashes; compare canonical relative paths.
    evidence = {Path(rel).as_posix(): entry for rel, entry in data['evidence_files'].items()}
    files = {p.relative_to(folder).as_posix(): p for p in folder.rglob('*')
             if p.is_file() and p != manifest}
    assert set(files) == set(evidence), f'Evidence file set drift: {folder}'
    for rel, expected in evidence.items():
        path = files[rel]
        assert sha(path) == expected['sha256'], f'Evidence hash drift: {path}'
        assert path.stat().st_size == expected['bytes'], f'Evidence size drift: {path}'
    return len(evidence), sha(manifest)


preflight = read(OUT / 'preflight.json')
assert preflight['status'] == 'PASS'
before = read(OUT / 'protected_before_rereview.json')
after = {path: sha(path) for path in before}
assert before == after, 'Protected source or previous evidence changed during rereview'
assert len(after) == 597
write(OUT / 'protected_after_rereview.json', after)

fix_count, fix_manifest_sha = verify_evidence(FIX)
previous_count, previous_manifest_sha = verify_evidence(PREVIOUS)
assert fix_count == 1111 and previous_count == 519
assert fix_manifest_sha == preflight['fix_manifest_sha']
assert previous_manifest_sha == '1cbf352694e064779a7f28234c6559d536f10091eb4f8a1e0908ec367c68aa58'
fix_manifest = read(FIX / 'REVIEW_MANIFEST.json')
product = {path: sha(path) for path in preflight['product_current']}
assert product == preflight['product_current']
assert product == fix_manifest['product_before_final_tests']
assert product == fix_manifest['product_after'] == fix_manifest['product_current']
assert product == read(OUT / 'provided_tests/product_before_final_tests.json')
assert product == read(OUT / 'provided_tests/product_after_final_tests.json')
assert product == read(OUT / 'independent_summary.json')['source_product_hashes']
write(OUT / 'product_current_final.json', product)

provided = read(OUT / 'provided_tests/targeted_tests.json')
assert provided['tests_run'] == 17 and provided['successful']
assert not any(provided[k] for k in ('failures', 'errors', 'skipped', 'network_attempts'))
assert provided['fixture_docx_generation'] == 0
independent = read(OUT / 'independent_cases.json')
assert len(independent) == 17 and all(case['status'] == 'PASS' for case in independent)
own_guard = read(OUT / 'independent_access_guard.json')
provided_guard = read(OUT / 'provided_tests/access_guard.json')
assert own_guard['production_reads'] == own_guard['writes_outside_rereview'] == 0
assert provided_guard['production_reads'] == provided_guard['writes_outside_run'] == 0
no_uzk = independent[0]['detail']
assert no_uzk['call_order'] == ['PROJECT', 'VOM', 'REGISTRY']
assert no_uzk['status'] == 'NO_UZK_REQUIRED' and no_uzk['GUI_generate_disabled']
assert no_uzk['GUI_message'] == 'Для текущего паспорта УЗК не требуется'
for field in ('marks', 'controlled_units', 'new_DOCX', 'new_directories', 'renderer_calls',
              'registry_mutations', 'profile_mapping_calls', 'Stage4_calls'):
    assert no_uzk[field] == 0

actual = read(OUT / 'actual_four_mark_preview.json')['data']
expected = {
    'Бп2-1': (8, [('1280', '5'), ('1024', '8.4')]),
    'Пс1-9': (9, [('3600', '4')]),
    'Пс1-10': (15, [('6000', '4')]),
    'Пс2-16': (1, [('480', '3')]),
}
observed = {}
for mark in actual['marks']:
    observed[mark['mark']] = (mark['Q'], [(str(s['L']), str(s['S'])) for s in mark['segments']])
    assert all(s['frequency_mhz'] == 5 and s['angle_degrees'] == 70 for s in mark['segments'])
assert observed == expected and actual['controlled_units'] == 33 and len(actual['marks']) == 4
assert actual['network_attempts'] == []
for path, expected_sha in actual['fingerprints'].items():
    assert sha(path) == expected_sha, f'Preview input/product drift: {path}'

docx = preflight['existing_docx']
assert Path(docx['path']).is_file()
assert sha(docx['path']) == docx['sha256'] == '7723be279f6e776a472fb33f0b1bc153aecd14fc04e5b319fae73f89f2960661'
assert product[str(ROOT / 'uzk_stage5/word.py')] == '61fad8e575023ec6c3e58096db180b44862ff715705988c59ae07b448a0c8fb0'

input_paths = [ROOT / 'DEV/ВЗМК_DEV.xlsx', ROOT / 'DEV/УЗК_DEV.xlsm',
               ROOT / 'DEV/Справочники/Профили.xlsx', Path(actual['mark_catalog']),
               Path(actual['vom']['path']), Path(actual['template'])]
inputs = {}
for path in input_paths:
    identity = str(path.resolve())
    inputs[identity] = {'before': before[identity], 'after': sha(path), 'unchanged': True}
    assert inputs[identity]['before'] == inputs[identity]['after']
write(OUT / 'input_hashes_before_after.json', inputs)

source_proof = {
    'assessment': 'S5-R1 CLOSED; no directly caused regression found',
    'changed_functions': preflight['changed_functions'],
    'patch_reverse_bound_to_previous_sha': preflight['reconstructed_source'],
    'ordering': [
        {'location': 'uzk_stage5/core.py:81-95', 'proof': 'Verify checkpoint, select passport/project, read and validate current visible VOM and shipment mass.'},
        {'location': 'uzk_stage5/inputs.py:153-178', 'proof': 'Read mandatory UZK registry; validate sheet/header/identities/positive quantities/duplicate consistency before use.'},
        {'location': 'uzk_stage5/core.py:96-98', 'proof': 'Only after registry validation compute normalized PROJECT+MARK intersection with visible VOM.'},
        {'location': 'uzk_stage5/core.py:100-103', 'proof': 'Skip mark_profiles and load_catalog entirely for empty intersection; calculate_marks excludes every mark without registry identity.'},
        {'location': 'uzk_stage5/core.py:34-71', 'proof': 'Unchanged quantity and rule calculation: per-mark shipment aggregation, registry membership, selected profile validation, Q=min(shipment,remaining). Both quantities are positive input-validated values.'},
        {'location': 'uzk_stage5/inputs.py:181-204', 'proof': 'Normalize requested mark set and skip unrelated rows before text/profile-conflict checks. Default required_marks=None remains strict.'},
        {'location': 'uzk_stage5/core.py:106-111', 'proof': 'Explicit NO_UZK_REQUIRED iff marks empty; controlled_units is sum of Q.'},
        {'location': 'uzk_stage5/core.py:119-141', 'proof': 'Unchanged generate verifies plan and pins, then returns None for empty marks before output-path creation, mkdir or renderer call.'},
        {'location': 'uzk_stage5/gui.py:150-155', 'proof': 'Exact no-UZK message in preview and generate completion; update_buttons disables generation for empty marks.'}
    ],
    'scope': 'Three modified runtime functions only, new targeted test; calculation/generate/Word renderer unchanged.',
    'caveat': 'General source fingerprinting/checkpoint safety still reads hashes of profile inputs even on no-UZK. Semantic profile mappings and Stage4 catalog lookup are skipped.'
}
source_proof['reviewed_functions'] = {}
for filename, names in [('uzk_stage5/core.py', {'preview', 'calculate_marks', 'generate'}),
                        ('uzk_stage5/inputs.py', {'registry_rows', 'mark_profiles'}),
                        ('uzk_stage5/gui.py', {'poll', 'update_buttons'})]:
    source = (ROOT / filename).read_text(encoding='utf-8-sig')
    lines = source.splitlines()
    for node in ast.walk(ast.parse(source)):
        if isinstance(node, ast.FunctionDef) and node.name in names:
            source_proof['reviewed_functions'][f'{filename}:{node.lineno}'] = {
                'symbol': node.name, 'end_line': node.end_lineno,
                'source': '\n'.join(lines[node.lineno - 1:node.end_lineno])}
write(OUT / 'source_order_assessment.json', source_proof)

not_run = [
    'Full Stage5 regression and full repeated Word review',
    'New real DEV DOCX generation; existing TEST DOCX modification',
    'Native Word visual rerun (previous PASS reused; renderer and accepted DOCX unchanged)',
    'Manual USER E2E; GUI state was checked programmatically in isolated Tk fixture',
    'Stage1-4 workflows, registry write/decrement, passport close, production access',
    'Stage6, PDF/EXE/deploy'
]
integrity = {
    'status': 'PASS', 'protected_before_after': {'count': len(after), 'changes': []},
    'product_BEFORE_AFTER_CURRENT_match': True, 'product_count': len(product),
    'fix_evidence_files_verified': fix_count, 'fix_manifest_sha256': fix_manifest_sha,
    'previous_evidence_files_verified': previous_count,
    'previous_manifest_sha256': previous_manifest_sha,
    'existing_docx': docx, 'existing_docx_unchanged': True,
    'word_renderer_unchanged': True, 'input_hashes_before_after': inputs,
    'provided_targeted_tests': provided, 'independent_tests': {'tests': 17, 'passed': 17, 'failures': []},
    'production_reads': 0, 'product_mutations': 0, 'registry_mutations': 0,
    'new_generated_DOCX': 0, 'no_UZK_generated_output_directories': 0,
    'scope_of_zero_outputs': 'Generate operations; isolated fixture/template copies are test inputs, not generated acts.',
    'previous_work_review_unchanged': True, 'not_run': not_run
}
write(OUT / 'integrity.json', integrity)
write(OUT / 'FINDINGS.json', {'S5-R1': 'CLOSED', 'BLOCKER': 0, 'MAJOR': 0, 'MINOR': 0, 'findings': [], 'required_fix': 'NONE'})

negative_lines = '\n'.join(f"   PASS | {case['name']} | {case['detail'] if isinstance(case['detail'], str) else json.dumps(case['detail'], ensure_ascii=False)}"
                           for case in independent[1:-1])
product_lines = '\n'.join(f'   {Path(path).relative_to(ROOT)} | {digest}' for path, digest in product.items())
input_lines = '\n'.join(f'   {path}\n   BEFORE=AFTER {item["after"]}' for path, item in inputs.items())
report = f'''ВЗМК — ЭТАП 5 TARGETED INDEPENDENT REREVIEW — S5-R1
Дата: {datetime.now(timezone.utc).isoformat()}
Scope: только S5-R1, его fix и evidence; продукт не исправлялся.

1. VERDICT — PASS
   S5-R1 CLOSED. Непосредственно вызванных исправлением регрессий не обнаружено.

2. CURRENT ARTIFACT VERIFIED — YES
   Фактические 11 product files совпали с fix tested BEFORE/AFTER/CURRENT,
   независимым повторным запуском targeted suite и final CURRENT.
   Exact changed runtime set: uzk_stage5/core.py, inputs.py, gui.py.
   Added: tests/test_stage5_s5_r1.py. Прочих изменений продукта нет.
   Patch независимо обращён к предыдущим SHA; AST: только preview, mark_profiles, App.poll.

3. S5-R1 ROOT CAUSE ASSESSMENT — CLOSED
   Ранее проверялся весь Справочник марок до определения потребности.
   Сейчас: проект/паспорт → валидный видимый ВОМ/масса → валидный обязательный реестр
   → normalized PROJECT+MARK intersection → только нужные mappings/Stage4.
   При пустом пересечении mappings и Stage4 catalog/lookup не вызываются.
   Q=min(суммарное количество марки в ВОМ, остаток реестра) не изменён.
   Для совпавших марок оба количества положительны после валидации; формулы не менялись.
   Fingerprinting/checkpoint protection остаётся обязательной и читает SHA входов.
   Точные строки исходников и доказательство последовательности: source_order_assessment.json.

4. NO_UZK_REQUIRED PREVIEW — PASS
   Независимо создана новая isolated fixture: валидный ВОМ, валидный реестр,
   ноль совпадений PROJECT+MARK, две посторонние конфликтующие марки и ПРОВЕРИТЬ.
   status=NO_UZK_REQUIRED; 0 марок; 0 изделий; MARK_PROFILE_CONFLICT отсутствует.
   Наблюдаемый порядок: PROJECT → VOM → REGISTRY; mappings=0; Stage4=0.

5. PREVIEW → GENERATE ASSESSMENT — PASS
   Реальный generate(plan, explicit_output) и generate(plan) вернули None.
   Новых DOCX=0; output directories=0; renderer calls=0; registry mutations=0.
   Файлы и список каталогов fixture до/после идентичны.
   GUI: «Для текущего паспорта УЗК не требуется»; кнопка генерации disabled,
   в том числе при установленном флажке подтверждения. Проверено на actual Tk App.poll.
   Копии шаблона внутри fixture являются входами тестов, не результатами Generate.

6. NEGATIVE CASES — A–H PASS
{negative_lines}
   Некорректный ВОМ останавливается до чтения registry_rows; отсутствие реестра
   блокируется checkpoint guard. Проверки входов не обходились.
   Для AMBIGUOUS/CONFLICT/отсутствия approved rule использованы isolated Catalog
   с принятой Stage4 provenance; реальные DEV sidecars не изменялись.

7. FOUR-MARK REGRESSION — PASS; только read-only preview
   Проект Гараж Казахстан; паспорт 9/2026.
   Бп2-1 Q=8: 1280×5, 1024×8.4.
   Пс1-9 Q=9: 3600×4.
   Пс1-10 Q=15: 6000×4.
   Пс2-16 Q=1: 480×3.
   Всего 4 отдельные марки, 33 изделия; все участки 5 МГц / 70°.
   ВОМ: 104 видимые строки; масса 15822.14 кг. DOCX не создавался.
   Evidence: actual_four_mark_preview.json.

8. EXISTING TEST DOCX INTEGRITY — PASS
   Exact path: {docx['path']}
   SHA-256: {docx['sha256']}
   Файл существует и неизменён. Word renderer SHA неизменён:
   {product[str(ROOT / 'uzk_stage5/word.py')]}
   calculate_marks, generate и Word renderer не изменены; предыдущий native Word PASS
   сохранён. Причины для повторной генерации или visual review не выявлены.

9. PRODUCT / INPUT SAFETY — PASS
   {len(after)} защищённых файлов BEFORE=AFTER: продукт, DEV, входы и previous review.
   Production ВЗМК.xlsx / УЗК.xlsm не читались, не хешировались, не изменялись.
   Все mutations — только в новом rereview evidence folder.
   Реестры/количества/паспорта не изменялись. Hash table: input_hashes_before_after.json.

10. EVIDENCE INTEGRITY — PASS
   Fix: independently verified all {fix_count} listed files by SHA/size/exact file set.
   Previous Work review: all {previous_count} evidence files + manifest unchanged.
   Provided targeted tests независимо перезапущены: 17/17 PASS, 0 skipped/errors/failures.
   Новые независимые сценарии: 17/17 PASS; actual current product используется напрямую.
   New REVIEW_MANIFEST.json записывается последним; после записи — только read-only
   SHA/size/exact file set/mtime verification. Его SHA возвращается finalizer stdout.
   Fix manifest SHA: {fix_manifest_sha}
   Previous Work manifest SHA: {previous_manifest_sha}

11. FINDINGS
   BLOCKER 0 / MAJOR 0 / MINOR 0.
   S5-R1 CLOSED. New findings: NONE.

12. NOT_RUN
   Полный Stage5 regression; повторный полный Word/native visual review;
   новая генерация реального DEV DOCX; ручной USER E2E;
   Stage1–4 workflows; production операции; изменение/списание реестра;
   закрытие паспорта; Stage6; PDF/EXE/deploy.

13. REQUIRED FIX — NONE

14. READY FOR USER E2E — YES
   Это готовность к отдельному USER E2E, не утверждение о выполненном ручном E2E.
   Проверка закончена. Продукт не исправлялся.

PRODUCT SHA-256 — fix tested BEFORE = AFTER = CURRENT = rereview final CURRENT:
{product_lines}

PROTECTED INPUT SHA-256:
{input_lines}

Evidence entry points:
   preflight.json; source_order_assessment.json; independent_cases.json;
   fresh_no_uzk_preview.json; actual_four_mark_preview.json;
   provided_tests/targeted_tests.json; provided_tests/targeted_tests.log;
   integrity.json; FINDINGS.json; REVIEW_MANIFEST.json.
'''
(OUT / 'REPORT.txt').write_text(report, encoding='utf-8')

# Final evidence write. No files are written after this operation.
evidence = {p.relative_to(OUT).as_posix(): {'sha256': sha(p), 'bytes': p.stat().st_size}
            for p in sorted(OUT.rglob('*')) if p.is_file() and p != MANIFEST}
new_manifest = {
    'created_utc': datetime.now(timezone.utc).isoformat(),
    'review': 'S5-R1 targeted independent rereview', 'status': 'PASS',
    'finding': 'S5-R1 CLOSED', 'findings': {'BLOCKER': 0, 'MAJOR': 0, 'MINOR': 0},
    'required_fix': 'NONE', 'ready_for_user_e2e': 'YES',
    'product_before': preflight['product_current'], 'product_current': product,
    'protected_files_count': len(after), 'protected_files_unchanged': True,
    'fix_manifest_sha256': fix_manifest_sha, 'fix_evidence_count': fix_count,
    'previous_manifest_sha256': previous_manifest_sha, 'previous_evidence_count': previous_count,
    'existing_test_docx': docx, 'no_UZK_required': no_uzk,
    'provided_tests': provided, 'independent_tests': {'tests': 17, 'passed': 17, 'failures': []},
    'four_mark_regression': {'marks': 4, 'units': 33, 'frequency_mhz': 5, 'angle_degrees': 70},
    'not_run': not_run, 'evidence_files': evidence, 'manifest_written_last': True,
}
write(MANIFEST, new_manifest)
count, manifest_sha = verify_evidence(OUT)
assert count == len(evidence)
assert MANIFEST.stat().st_mtime_ns >= max((OUT / rel).stat().st_mtime_ns for rel in evidence)
assert {path: sha(path) for path in before} == before
assert {path: sha(path) for path in product} == product
print(json.dumps({'status': 'PASS', 'new_manifest_verified_files': count,
                  'manifest_written_last_verified': True, 'manifest_sha256': manifest_sha,
                  'provided_tests': '17/17 PASS', 'independent_tests': '17/17 PASS',
                  'protected_unchanged': len(after), 'report': str(OUT / 'REPORT.txt')}, ensure_ascii=False, indent=2))
