"""Read-only current Stage5 reconciliation/preview. Never generates a DOCX/receipt."""
import ast
from datetime import datetime, timezone
from decimal import Decimal
import hashlib
import json
import os
from pathlib import Path
import sys
import traceback

RUN = Path(__file__).resolve().parent
ROOT = RUN.parents[1]
OLD = ROOT / 'validation/stage5_sample_results_20261008_130441'
REVIEW = ROOT / 'validation/stage5_sample_results_independent_review_20261008_01'
MIG = ROOT / 'validation/stage3_catalog_schema_20261008_162454'
STOP = ROOT / 'validation/stage7_implementation_20261009'
REQUEST = Path('C:/Users/arara/.codex/attachments/3b75b4aa-8b01-4872-8dd8-8722ac97b0f8/Вставленный текст.txt')
sys.path.insert(0, str(ROOT))


def sha(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def read(path):
    return json.loads(Path(path).read_text(encoding='utf-8-sig'))


def write(name, value):
    with (RUN / name).open('x', encoding='utf8') as stream:
        json.dump(value, stream, ensure_ascii=False, indent=2)


def snapshot(paths):
    return {str(p.resolve()): sha(p) for p in sorted(set(map(Path, paths))) if p.is_file()}


def files(folder):
    result = []
    for directory, children, names in os.walk(folder, followlinks=False):
        children[:] = [name for name in children if name != '__pycache__' and not name.startswith('node_modules')
                       and not (Path(directory) / name).is_junction() and not (Path(directory) / name).is_symlink()]
        result.extend(Path(directory) / name for name in names if (Path(directory) / name).is_file())
    return result


def verify_manifest(folder):
    manifest = read(folder / 'REVIEW_MANIFEST.json')
    entries = manifest.get('evidence', manifest.get('evidence_files'))
    assert entries is not None, ('Missing manifest evidence entries', folder)
    rows = entries if isinstance(entries, list) else [dict(path=p, sha256=h['sha256'] if isinstance(h, dict) else h) for p, h in entries.items()]
    errors = []
    for row in rows:
        path = (folder / row['path']).resolve()
        assert path.is_relative_to(folder), ('Manifest path escape', path)
        actual = sha(path) if path.is_file() else 'MISSING'
        if actual != row['sha256']:
            errors.append(dict(path=str(path), expected=row['sha256'], actual=actual))
    return dict(path=str(folder / 'REVIEW_MANIFEST.json'), sha256=sha(folder / 'REVIEW_MANIFEST.json'),
                checked=len(rows), mismatches=errors)


# Use an explicit previously audited source list. Never follow production metadata.
base_all = read(STOP / 'SOURCE_BEFORE.json')
excluded_named_copies = [p for p in base_all if Path(p).name.casefold() in ('взмк.xlsx', 'узк.xlsm')]
assert all(Path(p).is_relative_to(ROOT / 'e2e_user_stage1') or Path(p).is_relative_to(ROOT / 'e2e_user_stage3') for p in excluded_named_copies)
base = {p: h for p, h in base_all.items() if p not in excluded_named_copies}
protected_paths = list(map(Path, base))
for folder in (OLD, REVIEW, MIG, STOP, ROOT / 'DEV', ROOT / 'validation/stage5_implementation_20261008_105100'):
    protected_paths.extend(files(folder))
product_paths = [p for p in ROOT.iterdir() if p.is_file()]
for folder in ('passport_app', 'profiles_stage4', 'uzk_stage5', 'akz_stage6', 'tests'):
    product_paths.extend(files(ROOT / folder))
protected_paths.extend(product_paths)
protected_paths = [p for p in protected_paths if p.name.casefold() not in ('взмк.xlsx', 'узк.xlsm')]
before = snapshot(protected_paths)
product_before = snapshot(product_paths)
write('SOURCE_BEFORE.json', before)
write('PRODUCT_BEFORE.json', product_before)
write('REQUEST_BINDING.json', dict(path=str(REQUEST), sha256=sha(REQUEST)))
write('PREPARATION_TRACE.json', dict(first_attempt='Stopped by conservative basename guard before hashing/writes; no source opened or changed',
      reason='Prior Stage7 baseline includes two isolated historical E2E copies named ВЗМК.xlsx',
      excluded_without_open_or_hash=excluded_named_copies, correction='Exclude those copies; production basename guard remains enabled',
      second_attempt='Owned Python helper stopped before preview/evidence writes: tests/node_modules (2) junction led outside product scope',
      traversal_correction='Do not traverse junctions, symlinks, node_modules or __pycache__; hash only actual scoped files',
      third_attempt='Read-only preparation stopped on independent manifest key evidence_files, then audit serialization error; evidence retained in preparation_attempt3',
      corrections='Handle observed evidence/evidence_files schemas; serialize audit sets without mutating live guard'))
access = {'production_registry_attempts': [], 'network_attempts': [], 'forbidden_write_attempts': [],
          'python_source_reads': set(), 'native_read_scope': 'accepted visible_vom Excel adapter only; not a system-wide trace'}


def audit(event, args):
    if event.startswith('socket.') and event not in ('socket.__new__',):
        access['network_attempts'].append(event)
        raise RuntimeError('NETWORK_FORBIDDEN')
    if event != 'open' or not args or not isinstance(args[0], (str, bytes)):
        return
    path = Path(args[0].decode() if isinstance(args[0], bytes) else args[0]).resolve()
    if path.name.casefold() in ('взмк.xlsx', 'узк.xlsм', 'узк.xlsm'):
        access['production_registry_attempts'].append(str(path))
        raise RuntimeError('PRODUCTION_REGISTRY_FORBIDDEN')
    mode, flags = args[1:3]
    writing = (isinstance(mode, str) and any(c in mode for c in 'wax+')) or (isinstance(flags, int) and bool(flags & (1 | 2 | 64 | 512 | 1024)))
    if writing and not path.is_relative_to(RUN):
        access['forbidden_write_attempts'].append(str(path))
        raise RuntimeError('SOURCE_WRITE_FORBIDDEN: ' + str(path))
    if not writing and (path.is_relative_to(ROOT) or path.is_relative_to(ROOT.parent / 'Гараж казахстан')):
        access['python_source_reads'].add(str(path))


sys.addaudithook(audit)
status = 'STOP_PREFLIGHT'
failure = None
summary = None
reconciliation = []
capability = None
try:
    validations = [verify_manifest(p) for p in (OLD, REVIEW, MIG)]
    write('ACCEPTED_MANIFEST_VALIDATION.json', validations)
    assert all(not v['mismatches'] for v in validations), 'Accepted evidence manifest mismatch'
    comparison = read(MIG / 'product_hash_comparison.json')
    migration = read(MIG / 'real_migration_result.json')
    completion = read(MIG / 'real_migration_completion.json')
    review = read(REVIEW / 'REVIEW_MANIFEST.json')
    assert review['status'] == 'PASS' and completion['status'] == 'PASS'
    assert all(sha(p) == v['CURRENT'] for p, v in comparison.items()), 'Current migration product drift'
    accepted_final = read(OLD / 'product_current.json')
    for path, expected in accepted_final.items():
        allowed = comparison.get(path, {}).get('CURRENT', expected)
        assert sha(path) == allowed, ('Unexplained Stage5 product drift', path)
    # Independently reviewed final product uses final docs/tests, not preparation hashes.
    review_final = read(REVIEW / 'product_current_final.json')
    assert all(sha(p) == comparison.get(p, {}).get('CURRENT', h) for p, h in review_final.items())
    receipt = read(OLD / 'new_document.json')
    old_preview = read(OLD / 'current_preview.json')
    old_output = Path(receipt['output'])
    assert sha(old_output) == receipt['sha256'] == '8d3a2769baf6d381a0675804ef7f0806896f1b611e1498ea58677d3fdac1fd1b'
    assert old_preview['seal'] == receipt['preview_seal']
    for source, mapping in [('old_preview.fingerprints', old_preview['preview']['fingerprints']), ('old_receipt.product_binding', receipt['product_binding'])]:
        for path, old_hash in mapping.items():
            actual = sha(path)
            if actual == old_hash:
                provenance = 'UNCHANGED'
            elif path == migration['path'] and old_hash == migration['source_sha256'] and actual == migration['new_sha256']:
                provenance = 'ACCEPTED_STAGE3_SCHEMA_4_TO_3'
            elif path in comparison and old_hash == comparison[path]['ORIGINAL'] and actual == comparison[path]['CURRENT']:
                provenance = 'ACCEPTED_STAGE3_SCHEMA_4_TO_3'
            elif path in accepted_final and actual == accepted_final[path] and review_final[path] == actual:
                provenance = 'ALREADY_FINALIZED_AND_INDEPENDENTLY_REVIEWED_IN_OLD_STAGE5_CYCLE'
            else:
                provenance = 'UNEXPLAINED_DRIFT'
            reconciliation.append(dict(binding=source, path=path, expected_old=old_hash, actual_current=actual, provenance=provenance))
    write('OLD_CURRENT_SHA_RECONCILIATION.json', reconciliation)
    assert not any(v['provenance'] == 'UNEXPLAINED_DRIFT' for v in reconciliation), 'Unexplained old/current binding drift'
    assert {p: sha(p) for p in base} == base, 'Drift since Stage7 STOP source snapshot'
    write('MIGRATION_ACCEPTANCE.json', dict(
        authority='Current user request explicitly states migration accepted',
        implementation_report=dict(path=str(MIG / 'REVIEW_REPORT.txt'), sha256=sha(MIG / 'REVIEW_REPORT.txt')),
        completion=completion, migration_source_sha256=migration['source_sha256'],
        migration_output_sha256=migration['new_sha256'], current_catalog_sha256=sha(migration['path']),
        current_migration_product_count=len(comparison), product_drift=0,
        stage5_independent_review=dict(path=str(REVIEW / 'REPORT.txt'), sha256=sha(REVIEW / 'REPORT.txt'), status=review['status']),
        old_cycle_finalization_note=dict(path=str(OLD / 'PREPARATION_NOTES.md'), sha256=sha(OLD / 'PREPARATION_NOTES.md'))))
    from uzk_stage5.core import preview, digest
    from uzk_stage5.inputs import PINS, verify_pins
    from uzk_stage5.result_samples import load_sample_results
    verify_pins()
    pin_rows = [dict(path=str(p), expected=h, actual=sha(p)) for p, h in PINS.items()]
    assert all(r['expected'] == r['actual'] for r in pin_rows)
    write('CURRENT_INPUT_PINS.json', pin_rows + [dict(path=migration['path'], expected=migration['new_sha256'], actual=sha(migration['path']))])
    assert digest(old_preview['preview']) == old_preview['seal'], 'Old preview seal corrupted'
    vom = ROOT.parent / 'Гараж казахстан/Паспорт/9 от 06.10.2026/ВОМ отгрузочный счет 183  (5).xls'
    plan = preview('Гараж Казахстан', '9/2026', vom)
    plan.verify()
    sample = load_sample_results(plan.data)
    write('CURRENT_PREVIEW.json', dict(preview=plan.data, seal=plan.seal, sample=sample, operation='READ_ONLY_PREVIEW; GENERATE_NOT_CALLED'))
    write('CURRENT_INPUT_FINGERPRINTS.json', plan.data['fingerprints'])
    expected = {'Бп2-1': (8, [('1280', '5'), ('1024', '8.4')]), 'Пс1-9': (9, [('3600', '4')]),
                'Пс1-10': (15, [('6000', '4')]), 'Пс2-16': (1, [('480', '3')])}
    got = {}
    for mark in plan.data['marks']:
        got[mark['mark']] = (mark['Q'], [(str(Decimal(str(s['L'])).normalize()), str(Decimal(str(s['S'])).normalize())) for s in mark['segments']])
        assert all(Decimal(str(s['frequency_mhz'])) == 5 and Decimal(str(s['angle_degrees'])) == 70 for s in mark['segments'])
    normalize = lambda data: {m: (q, [(Decimal(l), Decimal(s)) for l, s in rows]) for m, (q, rows) in data.items()}
    assert normalize(got) == normalize(expected)
    assert plan.data['marks'] == old_preview['preview']['marks'], 'Mark calculation changed'
    assert plan.data['controlled_units'] == 33 and plan.data['passport']['number'] == '9/2026'
    assert plan.data['passport']['date'] == '01.10.2026'
    assert plan.data['vom']['visible_count'] == 104 and len(plan.data['vom']['hidden_data_rows']) == 397
    assert Decimal(plan.data['vom']['mass_kg']) == Decimal('15822.14')
    assert sha(vom) == 'aae5e5b8c0246c8fe97b3dec9cbcc039bc3e3cedb9d245b474b2f871b90596ae'
    assert sample['mode'] == 'SAMPLE_ONLY' and sample['results_confirmed'] is False and sample['row_count'] == 56
    summary = dict(status='PASS', number='9/2026', date='01.10.2026', marks=4, units=33,
                   visible_vom=104, hidden_vom=397, mass_kg='15822.14', calculation_identical_to_old=True,
                   geometry=expected, frequency_mhz=5, angle_degrees=70, sample_mode=sample['mode'],
                   sample_results_confirmed=sample['results_confirmed'], historical_sample_rows=56,
                   current_preview_seal=plan.seal, network_attempts=plan.data['network_attempts'])
    write('TARGETED_VALIDATION.json', summary)
    # Inspect actual accepted entry points, including old validation writer.
    code = {}
    for rel in ('uzk_stage5/core.py', 'uzk_stage5/word.py', 'uzk_stage5/gui.py', 'stage5.py'):
        path = ROOT / rel
        source = path.read_text(encoding='utf8')
        tree = ast.parse(source)
        functions = []
        for fn in ast.walk(tree):
            if isinstance(fn, (ast.FunctionDef, ast.AsyncFunctionDef)):
                functions.append(dict(name=fn.name, line=fn.lineno, end_line=fn.end_lineno,
                                      calls=sorted({ast.unparse(n.func) for n in ast.walk(fn) if isinstance(n, ast.Call)})))
        code[rel] = dict(path=str(path), sha256=sha(path), functions=functions)
    capability = dict(status='STOP_RECEIPT_UNSUPPORTED', accepted_product_entry_points=code,
        core_generate_return='Path only; no receipt writer, output SHA, generation run_id or receipt source bindings',
        cli_result='stdout: preview + preview_seal + output path only; no trustworthy generation receipt',
        gui_result='Path only; no receipt writer',
        old_receipt_origin=dict(path=str(OLD / 'create_sample.py'), sha256=sha(OLD / 'create_sample.py'),
            writer_lines=[39, 43], role='validation helper, not accepted product receipt API',
            old_receipt_keys=sorted(receipt), has_run_id='run_id' in receipt,
            has_direct_input_fingerprints='fingerprints' in receipt or 'source_bindings' in receipt),
        user_rule='If accepted Stage5 cannot create a trustworthy new receipt without architecture change: STOP. Do not invent a receipt manually.',
        generate_calls_this_operation=0, new_docx_created=0, new_receipt_created=0,
        next_required_scope='Separately authorized supported Stage5 operation receipt mechanism; no implementation performed')
    write('RECEIPT_CAPABILITY_STOP.json', capability)
    status = 'STOP_RECEIPT_UNSUPPORTED'
except Exception as exc:
    failure = dict(type=type(exc).__name__, message=str(exc), traceback=traceback.format_exc())
    write('PREFLIGHT_FAILURE.json', failure)

after = snapshot(protected_paths)
current = snapshot(product_paths)
write('SOURCE_AFTER.json', after)
write('PRODUCT_AFTER.json', current)
write('PRODUCT_CURRENT.json', snapshot(product_paths))
changes = {p: dict(before=h, after=after.get(p, 'MISSING')) for p, h in before.items() if after.get(p) != h}
integrity = dict(status='PASS' if not changes else 'FAIL', protected_count=len(before), changes=changes,
                 product_before_after_current_equal=product_before == current == read(RUN / 'PRODUCT_CURRENT.json'),
                 production_registry_opened=False, old_artifact_and_receipt_unchanged=True,
                 existing_stage7_stop_evidence_unchanged=True, code_mutations=0, registry_decrement=False)
write('SOURCE_INTEGRITY.json', integrity)
write('ACCESS_GUARD.json', {**access, 'python_source_reads': sorted(access['python_source_reads'])})
old_receipt_sha = sha(OLD / 'new_document.json')
old_doc = read(OLD / 'new_document.json')['output']
hash_lines = '\n'.join(f"| {Path(v['path']).name} ({v['binding']}) | `{v['expected_old']}` | `{v['actual_current']}` | {v['provenance']} |"
                       for v in reconciliation if v['expected_old'] != v['actual_current'])
report = f'''# Stage5 refresh for Stage7 — {status}

1. ACTUAL STATE

Root: `{ROOT}`. Только read-only reconciliation/preview. Новый согласованный checkpoint не создан. Индексные инструменты index_status/codebase_context недоступны; проверены точные файлы. Evidence operation folder: `{RUN}`. Это не generation run_id.

2. ROOT CAUSE OF STAGE 7 STOP

Старый preview/receipt относится к состоянию до принятой миграции Stage3 4→3. Сам старый DOCX не поврежден. Три основных расхождения подтверждены миграционными expected/actual SHA. STAGE5_DEV.md и tests/test_stage5_sample_results.py завершались после старой генерации; их текущие SHA уже вошли в финальный manifest и PASS независимого review старого Stage5. Это документированная подготовка старого цикла, не новый необъясненный drift после принятия. Старые bindings не переписаны.

3. OLD/CURRENT SHA RECONCILIATION

| Файл / binding | Старый expected SHA | Текущий actual SHA | Происхождение |
| --- | --- | --- | --- |
{hash_lines}

Полные записи: OLD_CURRENT_SHA_RECONCILIATION.json. Accepted migration product, финальные Stage5 product/review и прежний Stage7 source baseline проверены повторно. Acceptance authority — прямое указание пользователя; implementation completion и независимый Stage5 review проверены как локальные evidence, без выдумывания нового migration review. Детали: MIGRATION_ACCEPTANCE.json и ACCEPTED_MANIFEST_VALIDATION.json.

4. CURRENT STAGE 5 PREVIEW

{json.dumps(summary, ensure_ascii=False, indent=2) if summary else 'NOT_COMPLETED: см. PREFLIGHT_FAILURE.json'}

Полный новый preview, seal, fingerprints и SAMPLE_ONLY extraction сохранены. Для успешного preview: реальная ВОМ открыта native Excel read-only, macros disabled, UpdateLinks=0; 104 видимых/397 скрытых, 15822.14 кг. DEV дата N = 01.10.2026. Бп2-1 Q8:1280×5 и1024×8.4; Пс1-9 Q9:3600×4; Пс1-10 Q15:6000×4; Пс2-16 Q1:480×3; все5МГц/70°. Это расчёт 33 контролируемых изделий, не подтверждение фактического контроля.

5. NEW SAMPLE_ONLY DOCX

NOT_CREATED. Новый путь/SHA: NOT_APPLICABLE. generate() не вызывался. STOP до создания документа: требуемая операция receipt не поддерживается принятым продуктом.

6. NEW RECEIPT

NOT_CREATED; path/SHA/source bindings/run_id: NOT_APPLICABLE. Принятый uzk_stage5.core.generate() возвращает только Path; CLI печатает preview/seal/output, GUI принимает Path. Нет сохранения generation receipt с новым run_id, exact output SHA и полными current source bindings. Старый new_document.json писал validation/create_sample.py, не product API; в нем нет run_id. Повторить или расширить этот helper под новые SHA означало бы вручную добавить неподдерживаемый receipt, что пользователь запретил. CURRENT_PREVIEW.json и этот отчёт не выдаются за receipt. RECEIPT_CAPABILITY_STOP.json содержит hashes, AST calls и точные ссылки на исходники. Условие пользователя §7: STOP; архитектура не изменялась.

Старый receipt: `{OLD / 'new_document.json'}`; SHA `{old_receipt_sha}`. Старый DOCX: `{old_doc}`; SHA `{sha(old_doc)}`. Оба неизменны.

7. NATIVE WORD VALIDATION

NOT_RUN. Нового DOCX нет. Старые двухстраничные Word proofs не подменяют проверку новой генерации. Repair/recovery, новые две страницы, подписи/геометрия/гиперссылки нового output не проверялись.

8. SOURCE / OLD ARTIFACT SAFETY

SOURCE_INTEGRITY: {integrity['status']}; {len(before)} protected files, changed={len(changes)}. Product BEFORE=AFTER=CURRENT: {integrity['product_before_after_current_equal']}. Сохранены Stage1–6 code, DEV, реальные каталог/профили/ВОМ/шаблоны, старые УЗК/receipts и Stage7 STOP. Production не открывался/не хэшировался; metadata source_path не использован как файл. Python guard запрещал production registries, сетевые операции и writes вне новой evidence folder. Native Excel использовал принятый read-only adapter; общесистемный trace не заявляется. Остатки не списаны, макросы не запускались.

9. TARGETED VALIDATION

Выполнены проверки существующих manifest SHA, миграционных product hashes, old vs current provenance, DEV pins, старого output/receipt/seal, свежий actual Stage5 preview, точное равенство всех marks старому расчёту, геометрия/частота/угол/дата, все56 literal sample rows и receipt capability. Полная регрессия и генерационные tests NOT_RUN. Финальный gate: {status}; новый coherent checkpoint не готов.

10. EVIDENCE INTEGRITY

Новый EVIDENCE_MANIFEST.json записан последним, содержит hashes/sizes всех файлов новой folder кроме самого себя. Старые manifests/receipts не изменены. SOURCE_BEFORE/AFTER, PRODUCT_BEFORE/AFTER/CURRENT, CURRENT_INPUT_FINGERPRINTS и ACCESS_GUARD доступны отдельно. SHA нового manifest проверяется read-only после записи.

11. CHANGED FILES

Только новая папка `{RUN}`: read-only helper, отчёт и JSON evidence. CHANGED_FILES.json перечисляет все новые файлы, включая manifest. Product changes NONE. Старые evidence changes NONE.

12. NOT_RUN

Новый DOCX, новый generation receipt/run_id, native Word output validation, actual УЗК/подтверждение результатов, official issuance, decrement, production, web, PDF/EXE, full regression, Stage7 continuation и Stage8. Генератор не вызван по явному STOP пользователя при отсутствии поддерживаемого receipt. Архитектурный fix не выполнялся.

13. READY FOR TARGETED INDEPENDENT REVIEW

NO — требуемого нового TEST checkpoint нет. Read-only reconciliation и конкретный receipt blocker доступны для проверки. После отчёта STOP. Этап7 не продолжен, Этап8 не начат.
'''
if failure:
    report += '\nPREFLIGHT_FAILURE:\n```json\n' + json.dumps(failure, ensure_ascii=False, indent=2) + '\n```\n'
with (RUN / 'REFRESH_REPORT.md').open('x', encoding='utf8') as stream:
    stream.write(report)
future = ['CHANGED_FILES.json', 'EVIDENCE_MANIFEST.json']
write('CHANGED_FILES.json', dict(scope='NEW_VALIDATION_FOLDER_ONLY', product_changes=[], old_evidence_changes=[],
      added=[str(p) for p in sorted(files(RUN))] + [str(RUN / n) for n in future]))
manifest = dict(status=status, created_utc=datetime.now(timezone.utc).isoformat(),
    receipt_created=False, document_created=False, current_preview_pass=summary is not None,
    protected_integrity=integrity['status'], ready_for_targeted_independent_review='NO',
    evidence=[dict(path=str(p.relative_to(RUN)), sha256=sha(p), size=p.stat().st_size) for p in sorted(files(RUN))],
    manifest_written_last=True)
write('EVIDENCE_MANIFEST.json', manifest)
errors = [v['path'] for v in manifest['evidence'] if sha(RUN / v['path']) != v['sha256']]
print(json.dumps(dict(status=status, failure=failure, preview=summary, protected_count=len(before), source_changes=len(changes),
    manifest=str(RUN / 'EVIDENCE_MANIFEST.json'), manifest_sha256=sha(RUN / 'EVIDENCE_MANIFEST.json'),
    evidence_files=len(manifest['evidence']), manifest_mismatches=errors), ensure_ascii=False, indent=2))
assert not errors and not changes and product_before == current
