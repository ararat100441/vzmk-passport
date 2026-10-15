"""Finalize explicit STOP evidence only. No Stage5 preview or generation calls."""
import ast
from datetime import datetime, timezone
import hashlib
import json
from pathlib import Path

RUN = Path(__file__).resolve().parent
ROOT = RUN.parents[1]
OLD = ROOT / 'validation/stage5_sample_results_20261008_130441'
REVIEW = ROOT / 'validation/stage5_sample_results_independent_review_20261008_01'
MIG = ROOT / 'validation/stage3_catalog_schema_20261008_162454'
S6 = ROOT / 'validation/stage6_final_targeted_rereview_r2_20261009'

def sha(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()

def read(path):
    return json.loads(Path(path).read_text(encoding='utf-8-sig'))

def write(name, data):
    with (RUN / name).open('x', encoding='utf8') as stream:
        json.dump(data, stream, ensure_ascii=False, indent=2)

# Retain the preparation STOP and its manifest; create a complete final report.
prep = RUN / 'preparation_attempt4'
assert not prep.exists()
prep.mkdir()
keep = {'refresh_preflight.py', 'finalize_stop.py', 'SOURCE_BEFORE.json', 'PRODUCT_BEFORE.json',
        'REQUEST_BINDING.json', 'ACCEPTED_MANIFEST_VALIDATION.json'}
for path in list(RUN.iterdir()):
    if path.is_file() and path.name not in keep:
        destination = prep / path.name
        assert path.resolve().is_relative_to(RUN) and destination.resolve().is_relative_to(RUN)
        path.rename(destination)

comparison = read(MIG / 'product_hash_comparison.json')
drift = [dict(path=p, migration_expected=v['CURRENT'], actual_current=sha(p))
         for p, v in comparison.items() if sha(p) != v['CURRENT']]
assert len(drift) == 1 and Path(drift[0]['path']) == ROOT / 'passport_app/vom.py'
prior_drift = read(S6 / 'guarded_mismatches_on_start.json')
assert any(v['path'] == drift[0]['path'] and v['actual'] == drift[0]['actual_current'] for v in prior_drift)
drift[0].update(provenance='OUTSIDE_SCHEMA_MIGRATION; already declared expected parallel work in Stage6 review',
    prior_review=str(S6 / 'REVIEW_REPORT.txt'), prior_review_sha256=sha(S6 / 'REVIEW_REPORT.txt'),
    likely_original_change_report=str(ROOT / 'validation/targeted_M1_M2_20261005_231515/TARGETED_FIX_REPORT.txt'),
    original_report_limit='Report proves Connections exact-text change scope; same actual SHA is recorded in before/after. No reconstruction of an unrecorded transition is claimed.')
write('OUTSIDE_MIGRATION_DRIFT_STOP.json', dict(status='STOP', records=drift,
    user_rule='Если есть drift вне принятой миграции: STOP.',
    decision='Do not silently extend current refresh acceptance to another cycle; no preview/generate call'))

receipt = read(OLD / 'new_document.json')
old_preview = read(OLD / 'current_preview.json')
accepted = read(OLD / 'product_current.json')
independent = read(REVIEW / 'product_current_final.json')
migration = read(MIG / 'real_migration_result.json')
rows = []
for binding, mapping in [('preview.fingerprints', old_preview['preview']['fingerprints']), ('receipt.product_binding', receipt['product_binding'])]:
    for path, expected in mapping.items():
        actual = sha(path)
        if actual == expected:
            origin = 'UNCHANGED'
        elif path == migration['path'] and expected == migration['source_sha256'] and actual == migration['new_sha256']:
            origin = 'ACCEPTED_STAGE3_SCHEMA_4_TO_3'
        elif path in comparison and expected == comparison[path]['ORIGINAL'] and actual == comparison[path]['CURRENT']:
            origin = 'ACCEPTED_STAGE3_SCHEMA_4_TO_3'
        elif path in accepted and actual == accepted[path] == independent[path]:
            origin = 'OLD_STAGE5_FINALIZATION_AFTER_GENERATION; independently reviewed final bytes'
        else:
            origin = 'UNEXPLAINED'
        rows.append(dict(binding=binding, path=path, expected_old=expected, actual_current=actual, provenance=origin))
assert not any(v['provenance'] == 'UNEXPLAINED' for v in rows)
write('OLD_CURRENT_SHA_RECONCILIATION.json', dict(records=rows, preview_fingerprint_count=len(old_preview['preview']['fingerprints']),
    preview_differences=sum(v['binding']=='preview.fingerprints' and v['expected_old']!=v['actual_current'] for v in rows),
    receipt_product_differences=sum(v['binding']=='receipt.product_binding' and v['expected_old']!=v['actual_current'] for v in rows),
    unexpected_stage5_bindings=0, broader_migration_snapshot_drift=drift))
write('MIGRATION_ACCEPTANCE.json', dict(authority='Direct current user request: accepted Stage3 4→3 migration',
    source_sha256=migration['source_sha256'], output_sha256=migration['new_sha256'],
    actual_catalog_sha256=sha(migration['path']), completion=read(MIG/'real_migration_completion.json'),
    current_product_rows=len(comparison), current_product_mismatches=drift,
    stage5_final_review_status=read(REVIEW/'REVIEW_MANIFEST.json')['status'],
    finalization_note=dict(path=str(OLD/'PREPARATION_NOTES.md'),sha256=sha(OLD/'PREPARATION_NOTES.md'))))
pins = {
    ROOT/'DEV/ВЗМК_DEV.xlsx':'8e56f68518c70cade9efb09084dd67cd5d72e07e3fdc813b8515030e9cb98e80',
    ROOT/'DEV/УЗК_DEV.xlsm':'9f25a30139ab9e31b6da7248284c2ae6e9d4327f0d68251d5723293992662302',
    ROOT/'DEV/Справочники/Профили.xlsx':'7b420f6dead96408680a64d220c83e8980aeec1c38635e35b2e4723a98e3831b',
    Path(migration['path']):'7fa36f9ecc2ebb722aff89fdb94ffedf330608dc73c16264005e31e11f4b1587'}
pin_rows = [dict(path=str(p),expected=h,actual=sha(p)) for p,h in pins.items()]
assert all(v['expected'] == v['actual'] for v in pin_rows)
write('CURRENT_INPUT_FINGERPRINTS.json', dict(scope='Pinned inputs and old-bound current files; NOT_A_NEW_PREVIEW',
    pins=pin_rows, actual_old_bound_sources={p:sha(p) for p in old_preview['preview']['fingerprints']}))
assert sha(receipt['output']) == receipt['sha256'] == '8d3a2769baf6d381a0675804ef7f0806896f1b611e1498ea58677d3fdac1fd1b'
assert old_preview['seal'] == receipt['preview_seal']
write('OLD_ARTIFACT_BINDING.json',dict(docx=receipt['output'],sha256=sha(receipt['output']),
    receipt_path=str(OLD/'new_document.json'),receipt_sha256=sha(OLD/'new_document.json'),
    preview_path=str(OLD/'current_preview.json'),preview_sha256=sha(OLD/'current_preview.json'),
    seal=receipt['preview_seal'],mode=receipt['mode'],results_confirmed=receipt['results_confirmed']))
code = {}
for relative in ('uzk_stage5/core.py','uzk_stage5/word.py','uzk_stage5/gui.py','stage5.py'):
    path = ROOT/relative
    tree = ast.parse(path.read_text(encoding='utf8'))
    functions=[]
    for fn in ast.walk(tree):
        if isinstance(fn,ast.FunctionDef):
            functions.append(dict(name=fn.name,line=fn.lineno,end_line=fn.end_lineno,
                calls=sorted({ast.unparse(n.func) for n in ast.walk(fn) if isinstance(n,ast.Call)})))
    code[relative]=dict(path=str(path),sha256=sha(path),functions=functions)
write('RECEIPT_CAPABILITY_STOP.json',dict(status='UNSUPPORTED_BY_ACCEPTED_PRODUCT',entry_points=code,
    generate_contract='Returns Path only. No operation receipt writer, generation run_id, saved output SHA or full saved source bindings.',
    cli_contract='stdout preview/seal/output only',gui_contract='Path only',
    prior_writer=dict(path=str(OLD/'create_sample.py'),sha256=sha(OLD/'create_sample.py'),lines=[39,43],
        role='Validation helper, not product receipt API',receipt_keys=sorted(receipt),has_run_id='run_id' in receipt),
    user_rule='Если принятый Stage5 не способен создать достоверный новый receipt без изменения архитектуры: STOP. Не придумывать receipt вручную.',
    product_or_helper_generation_calls=0,new_docx=0,new_receipt=0))
write('CURRENT_PREVIEW_NOT_RUN.json',dict(status='NOT_RUN',reason='Explicit STOP on outside-migration product snapshot drift',
    required_marks={'Бп2-1':{'Q':8,'geometry':['1280×5','1024×8.4']},'Пс1-9':{'Q':9,'geometry':['3600×4']},
        'Пс1-10':{'Q':15,'geometry':['6000×4']},'Пс2-16':{'Q':1,'geometry':['480×3']}},
    required_vom_visible=104,required_vom_hidden=397,required_mass_kg='15822.14',required_units=33,
    required_date='01.10.2026',required_frequency_mhz=5,required_angle_degrees=70,
    old_preview_and_migration_preview='Historical evidence only; not substituted for a current calculation'))
before=read(RUN/'SOURCE_BEFORE.json')
after={p:sha(p) if Path(p).is_file() else 'MISSING' for p in before}
product_before=read(RUN/'PRODUCT_BEFORE.json')
product_current={p:sha(p) if Path(p).is_file() else 'MISSING' for p in product_before}
changes={p:dict(before=h,after=after[p]) for p,h in before.items() if h!=after[p]}
assert not changes and product_before==product_current
write('SOURCE_AFTER.json',after)
write('PRODUCT_AFTER.json',product_current)
write('PRODUCT_CURRENT.json',product_current)
write('SOURCE_INTEGRITY.json',dict(status='PASS',protected_count=len(before),changed_files=changes,
    product_before_after_current_equal=True,old_docx_receipt_unchanged=True,stage7_stop_unchanged=True,
    production_open_or_hash=False,registry_decrement=False,
    excluded_two_E2E_copies='Conservative basename guard: isolated historical copies named ВЗМК.xlsx not opened or hashed in this task'))
write('TARGETED_VALIDATION.json',dict(status='STOP',accepted_manifest_files_verified=565,accepted_manifest_mismatches=0,
    pins='4/4 PASS',old_preview_bindings='27 checked; 3 explained migration changes',
    old_receipt_product_bindings='3 differences explained: inputs.py migration; docs/test old-cycle finalization',
    migration_product_drift=drift,source_integrity='PASS',new_current_preview='NOT_RUN',generate='NOT_RUN',native_word='NOT_RUN'))
diff_table='\n'.join(f"| {Path(v['path']).name} / {v['binding']} | `{v['expected_old']}` | `{v['actual_current']}` | {v['provenance']} |"
    for v in rows if v['expected_old']!=v['actual_current'])
report=f'''# Targeted Stage5 refresh for Stage7 — STOP

1. ACTUAL STATE

Новый coherent SAMPLE_ONLY checkpoint не создан. Read-only reconciliation завершена; свежий preview и генерация не запускались после обнаружения drift вне названной миграции. Root: `{ROOT}`. Индексные инструменты недоступны; использованы точные локальные source/evidence файлы.

2. ROOT CAUSE OF STAGE 7 STOP

Старые source bindings предшествуют принятой Stage3 schema migration 4→3. Старый DOCX исправен. Три основных SHA расхождения объяснены этой миграцией. Новый refresh дополнительно остановлен по пользовательскому §1: product snapshot миграции и текущий passport_app/vom.py различаются. Это отдельный цикл, не изменение 4→3.

3. OLD/CURRENT SHA RECONCILIATION

| Файл / binding | Старый expected SHA | Текущий actual SHA | Происхождение |
| --- | --- | --- | --- |
{diff_table}

Все27 старых preview fingerprints проверены:3 различия миграции. В receipt product_binding3 различия: inputs.py и два файла подготовки Stage5. STAGE5_DEV.md и tests/test_stage5_sample_results.py уже имеют текущие SHA в финальном Stage5 manifest и PASS независимого review; это завершение старого цикла после генерации, не необъясненное runtime изменение.

Дополнительный drift вне миграции:
`{drift[0]['path']}`
expected `{drift[0]['migration_expected']}`
actual `{drift[0]['actual_current']}`.
Stage6 final review уже отмечал ровно это изменение как ожидаемую параллельную работу, с записью ответа пользователя «Да, изменения другой работы ожидаемы». Изменение не приписано миграции и не замалчивается. Текущий контракт требует STOP при drift вне указанной миграции; его область принятия автоматически не расширялась. Подробности OUTSIDE_MIGRATION_DRIFT_STOP.json и OLD_CURRENT_SHA_RECONCILIATION.json.

4. CURRENT STAGE 5 PREVIEW

NOT_RUN по §1 STOP. Четыре current pins совпали полностью: DEV ВЗМК, DEV УЗК, Профили и project catalog. CURRENT_INPUT_FINGERPRINTS.json содержит actual SHA, но не выдается за новый preview. Исторические 4 марки/33 изделия,104/397 строк и15822.14кг не подменяют свежий расчёт. Требуемые Q/геометрия/5МГц/70°/01.10.2026 сохранены как ожидаемый контракт в CURRENT_PREVIEW_NOT_RUN.json.

5. NEW SAMPLE_ONLY DOCX

NOT_CREATED. Новый exact path/SHA: NOT_APPLICABLE. generate() не вызывался. Старый DOCX `{receipt['output']}` сохранён; SHA `{sha(receipt['output'])}`.

6. NEW RECEIPT

NOT_CREATED; новый path/SHA/source bindings/run_id: NOT_APPLICABLE. Дополнительный независимый blocker §7: принятый core.generate() возвращает Path, CLI выводит preview/seal/output, GUI получает Path. Поддерживаемой записи operation receipt с run_id, exact output SHA и полными current bindings нет. Старый receipt записывал validation/create_sample.py; у него отсутствует run_id. Новый receipt вручную не придуман, старые SHA не заменены, архитектура не изменена. AST, line references и code SHA сохранены в RECEIPT_CAPABILITY_STOP.json.

Старый receipt `{OLD/'new_document.json'}`; SHA `{sha(OLD/'new_document.json')}`. Старый preview seal `{receipt['preview_seal']}` совпал с receipt. SAMPLE_ONLY/results_confirmed=false сохранены.

7. NATIVE WORD VALIDATION

NOT_RUN: нового DOCX нет. Старый Word proof не выдается за проверку новой генерации. Две страницы, no repair, подписи и отсутствие PDF hyperlinks нового output не заявлены.

8. SOURCE / OLD ARTIFACT SAFETY

PASS: {len(before)} защищённых файлов, changed=0. Product BEFORE=AFTER=CURRENT. Stage1–6 code, DEV registries, каталог, профили, ВОМ, шаблоны, старые outputs/receipts и Stage7 STOP сохранены. Production не открывался/не хэшировался, web/макросы/списание не выполнялись. Две изолированные E2E-копии с basename ВЗМК.xlsx исключены консервативным guard без открытия/хэширования. Junction tests/node_modules (2) исключён из финального source inventory.

9. TARGETED VALIDATION

565 evidence files по трём принятым manifests: SHA PASS,0 mismatches. Current pins4/4 PASS; old artifact/seal PASS; источник трёх основных расхождений установлен; документация/test preparation связаны с финальным Stage5 review. Broader migration product comparison обнаружил1 отдельное изменение vom.py — STOP. Статический анализ receipt capability подтвердил отсутствие требуемого API. Новые расчётные/генерационные tests не запускались; полнота будущего checkpoint не заявлена.

10. EVIDENCE INTEGRITY

EVIDENCE_MANIFEST.json записан последним и хэширует все новые evidence, включая сохранённые подготовительные попытки; сам себя не включает. Manifest проверен read-only после записи. Подготовительный helper сначала остановился на именах E2E-копий; затем обход runtime junction был прерван у собственного процесса до preview/evidence writes; далее исправлены чтение manifest evidence_files и сериализация audit set. Подготовительные журналы сохранены; product changes0. Старые manifests не менялись.

11. CHANGED FILES

Только новая папка `{RUN}`: два helper, reconciliation/pins/capability/source evidence, отчёт, сохранённая preparation evidence и manifest. CHANGED_FILES.json содержит полный список. Product changes NONE; старые evidence changes NONE.

12. NOT_RUN

Fresh Stage5 preview/native Excel, новая DOCX generation, operation receipt/run_id, native Word output validation, actual УЗК, подтверждение исполнителя, official issuance, inventory decrement, production, web, PDF/EXE, full regression, Stage7 continuation и Stage8. Нет нового generation run_id; имя evidence folder не используется как его замена.

13. READY FOR TARGETED INDEPENDENT REVIEW

NO — требуемого нового согласованного checkpoint нет. Reconciliation и два конкретных STOP основания доступны для проверки. После отчёта STOP. Stage7 не продолжен; Stage8 не начат.
'''
with (RUN/'REFRESH_REPORT.md').open('x',encoding='utf8') as stream:stream.write(report)
new_files=[p for p in RUN.rglob('*') if p.is_file()]
write('CHANGED_FILES.json',dict(product_changes=[],old_evidence_changes=[],
    added=[str(p) for p in sorted(new_files)]+[str(RUN/'CHANGED_FILES.json'),str(RUN/'EVIDENCE_MANIFEST.json')]))
manifest=dict(status='STOP_OUTSIDE_MIGRATION_DRIFT_AND_RECEIPT_UNSUPPORTED',created_utc=datetime.now(timezone.utc).isoformat(),
    current_preview='NOT_RUN',new_docx=False,new_receipt=False,ready_for_targeted_independent_review='NO',
    evidence=[dict(path=str(p.relative_to(RUN)),sha256=sha(p),size=p.stat().st_size) for p in sorted(RUN.rglob('*')) if p.is_file()],
    manifest_written_last=True)
write('EVIDENCE_MANIFEST.json',manifest)
errors=[v['path'] for v in manifest['evidence'] if sha(RUN/v['path'])!=v['sha256']]
assert not errors
print(json.dumps(dict(status=manifest['status'],protected=len(before),source_changes=0,
    evidence_files=len(manifest['evidence']),manifest_mismatches=errors,manifest_path=str(RUN/'EVIDENCE_MANIFEST.json'),
    manifest_sha256=sha(RUN/'EVIDENCE_MANIFEST.json')),ensure_ascii=False,indent=2))
