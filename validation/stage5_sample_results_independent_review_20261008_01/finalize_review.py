"""Read-only final checks; write only this review's evidence; manifest is last."""
import ast
import hashlib
import json
from datetime import datetime,timezone
from pathlib import Path

OUT=Path(__file__).resolve().parent
ROOT=OUT.parent.parent
IMPL=ROOT/'validation/stage5_sample_results_20261008_130441'
MANIFEST=OUT/'REVIEW_MANIFEST.json'

def sha(path):
    path=Path(path)
    assert path.name.casefold() not in ('взмк.xlsx','узк.xlsm')
    return hashlib.sha256(path.read_bytes()).hexdigest()
def read(path):return json.loads(Path(path).read_text(encoding='utf-8-sig'))
def save(name,data):
    path=(OUT/name).resolve();assert path.is_relative_to(OUT)
    path.write_text(json.dumps(data,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')

preflight=read(OUT/'preflight.json')
assert preflight['status']=='PASS'
protected=read(OUT/'protected_before.json')
after={p:sha(p) for p in protected}
assert after==protected and len(after)==2952
save('protected_after.json',after)
product_before=read(OUT/'product_before_review.json')
product={p:sha(p) for p in product_before}
assert product==product_before and len(product)==13
for name in ('provided_tests/product_before_tests.json','provided_tests/product_after_tests.json',
             's5r1_product_before.json','s5r1_product_after.json',
             'independent_product_before.json','independent_product_after.json'):
    assert read(OUT/name)==product,name
save('product_current_final.json',product)

impl=read(IMPL/'REVIEW_MANIFEST.json')
assert impl['product_current']==impl['product_after']==impl['test_binding_before']==impl['test_binding_after']==product
assert sha(IMPL/'REVIEW_MANIFEST.json')==preflight['implementation_manifest_sha256']
impl_files={Path(e['path']).as_posix():e for e in impl['evidence']}
actual_impl={p.relative_to(IMPL).as_posix():p for p in IMPL.rglob('*') if p.is_file() and p.name!='REVIEW_MANIFEST.json'}
assert set(actual_impl)==set(impl_files)
for rel,e in impl_files.items():
    assert sha(actual_impl[rel])==e['sha256'] and actual_impl[rel].stat().st_size==e['size']
assert len(impl_files)==281
assert (IMPL/'REVIEW_MANIFEST.json').stat().st_mtime_ns>=max(p.stat().st_mtime_ns for p in actual_impl.values())

provided=read(OUT/'provided_tests/targeted_tests.json')
s5r1=read(OUT/'s5r1_targeted.json')
independent=read(OUT/'independent_summary.json')
for d,n in ((provided,16),(s5r1,4)):
    assert d['tests_run']==n and d['successful'] and not any(d[k] for k in ('errors','failures','network_attempts'))
assert independent['tests']==independent['passed']==16 and independent['failures']==independent['network_attempts']==[]
for name,write_key in [('provided_tests/targeted_access_guard.json','writes_outside_validation'),
                       ('s5r1_access_guard.json','writes_outside_run'),
                       ('independent_access_guard.json','writes_outside_review')]:
    d=read(OUT/name)
    assert d['production_reads']==d[write_key]==0

empty=read(OUT/'B_no_output.json')
assert empty['status']=='NO_UZK_REQUIRED'
assert all(empty[k]==0 for k in ('new_docx','new_output_directories','source_mutations','writer_calls'))
assert s5r1['fixture_docx_generation']==0
fresh=read(OUT/'fresh_current_preview.json')['data']
assert len(fresh['marks'])==4 and fresh['controlled_units']==33
assert fresh['passport']['number']=='9/2026' and fresh['passport']['date']=='01.10.2026'
for p,h in fresh['fingerprints'].items():assert sha(p)==h
native=read(OUT/'native_word/native_word.json')
visual=read(OUT/'visual_validation_independent.json')
assert native['status']=='PASS_NORMAL_READONLY_OPEN' and native['readonly'] and native['unchanged']
assert native['close_without_save'] and native['existing_word_processes_preserved']
assert visual['status']=='PASS' and visual['pages']==2 and visual['landscape']
artifact=Path(preflight['artifact'])
assert sha(artifact)==preflight['artifact_sha256']=='8d3a2769baf6d381a0675804ef7f0806896f1b611e1498ea58677d3fdac1fd1b'
old=preflight['old_test_docx']
assert sha(old['path'])==old['sha256']=='7723be279f6e776a472fb33f0b1bc153aecd14fc04e5b319fae73f89f2960661'
history=read(OUT/'historical_rows_independent.json')
sources=read(OUT/'historical_sources_independent.json')
assert len(history)==56 and len(sources)==10 and {(r['graph7'],r['graph8']) for r in history}=={('ДНО','годен')}

inputs=[ROOT/'DEV/ВЗМК_DEV.xlsx',ROOT/'DEV/УЗК_DEV.xlsm',ROOT/'DEV/Справочники/Профили.xlsx',
        Path(fresh['mark_catalog']),Path(fresh['vom']['path']),Path(fresh['template']),
        Path(old['path']),artifact,*(Path(item['path']) for item in sources)]
input_hashes={}
for p in inputs:
    path=str(p.resolve());assert protected[path]==sha(p)
    input_hashes[path]={'before':protected[path],'after':sha(p),'unchanged':True}
save('input_artifact_hashes_before_after.json',input_hashes)

scope={'status':'PASS','modified':['uzk_stage5/word.py','uzk_stage5/gui.py','STAGE5_DEV.md'],
       'added':['uzk_stage5/result_samples.py','tests/test_stage5_sample_results.py'],
       'calculation_registry_modules_unchanged':True,'old_stage1_4_and_previous_evidence_unchanged':True,
       'runtime_diff_proof':preflight['accepted_source_recovery'],
       'source_assessment':[
           {'location':'uzk_stage5/result_samples.py:88','proof':'load_sample_results rejects any kind except TEST/DRAFT before source access. Historical examples stay SAMPLE_ONLY and confirmed=False.'},
           {'location':'uzk_stage5/result_samples.py:95','proof':'Exactly ten bounded source paths, accepted SHA values, expected physical row counts and one exact 56-row literal pair are required.'},
           {'location':'uzk_stage5/result_samples.py:78','proof':'Blank fields or multiple distinct raw result pairs STOP; no case/space normalization of outcomes.'},
           {'location':'uzk_stage5/word.py:52','proof':'Load sample before template processing. Only title plus cells with zero-based indexes 6 and 7 are new content changes.'},
           {'location':'uzk_stage5/word.py:118','proof':'Recheck sample hashes before exclusive output open; changed/missing sample cannot yield a document.'},
           {'location':'uzk_stage5/gui.py:29','proof':'Only the sample/unconfirmed banner changed. Generation stays TEST; quantity confirmation checkbox is not confirmation of inspection results.'},
           {'location':'STAGE5_DEV.md:21','proof':'Future official contract requires each mark actual results, responsible executor, confirmation time and inspection evidence. Official confirmation/issuance workflow remains unimplemented.'}
       ]}
scope['current_functions']={}
for filename,names in [('uzk_stage5/result_samples.py',{'load_sample_results','assess_rows','verify_sample_sources'}),
                       ('uzk_stage5/word.py',{'write_test_document'})]:
    text=(ROOT/filename).read_text(encoding='utf-8-sig');lines=text.splitlines()
    for node in ast.walk(ast.parse(text)):
        if isinstance(node,ast.FunctionDef) and node.name in names:
            scope['current_functions'][f'{filename}:{node.lineno}']={'symbol':node.name,'end_line':node.end_lineno,
                                                                  'source':'\n'.join(lines[node.lineno-1:node.end_lineno])}
save('source_scope_assessment.json',scope)
findings={'BLOCKER':0,'MAJOR':0,'MINOR':0,'findings':[],'required_fix':'NONE'}
save('FINDINGS.json',findings)
not_run=['Actual ultrasonic inspection or per-mark executor confirmation','Official generation/issuance',
         'New act generation, including repeat creation of the exact TEST DOCX',
         'Full Stage5 regression or Stage1-4 workflows','Manual USER E2E after this change',
         'Registry decrement, passport closure, production operations','Stage6, deploy']
integrity={'status':'PASS','product_files':13,'BEFORE_AFTER_CURRENT_equal':True,'protected_files':len(protected),
           'protected_changes':[],'implementation_evidence_files':281,
           'implementation_manifest_sha256':sha(IMPL/'REVIEW_MANIFEST.json'),
           'implementation_manifest_last_verified':True,'input_artifact_hashes':input_hashes,
           'targeted_tests':provided,'independent_tests':independent,'selected_S5R1':s5r1,
           'new_generated_DOCX':0,'no_UZK_output_directories':0,'registry_mutations':0,
           'production_reads':0,'product_mutations':0,
           'fixture_note':'Historical/template copies and modified DOCX sample fixtures are test inputs, never newly issued acts.',
           'not_run':not_run}
save('integrity.json',integrity)

preparation='''Resolved review-helper preparation only; no product or source changes.
1. Computer Use could inspect the existing Word window, but its modal Open dialog
   state was not retained for the next action. No document text was typed or saved.
   Native validation used Word's document object model (not UI Automation).
2. Explicit optional Missing arguments were rejected by the PowerShell COM binder
   before document opening. native_word_preparation_attempt1.json is retained.
   Standard Documents.Open(path,False,True,False) then succeeded with ReadOnly=True,
   OpenAndRepair=False API default, and close-without-save. Both native page images
   were independently inspected at original resolution. Exact SHA stayed unchanged.
3. The accepted guide's prior text snapshot is absent. Its before/after hashes and
   complete current contract were reviewed. Runtime word.py and GUI before bytes
   were recovered/bound to accepted SHA; the actual runtime diff is fully proven.
'''
(OUT/'PREPARATION_NOTES.txt').write_text(preparation,encoding='utf-8')
product_lines='\n'.join(f'   {Path(p).relative_to(ROOT)} | {h}' for p,h in product.items())
source_lines='\n'.join(f'   {item["path"]}\n   SHA-256 {item["sha256"]}; rows={item["rows"]}' for item in sources)
negative_lines='\n'.join('   '+r['name']+' — '+r['status'] for r in read(OUT/'independent_checks.json'))
report=f'''ВЗМК — ЭТАП 5 TARGETED INDEPENDENT REVIEW — автозаполнение результатов УЗК
Date UTC: {datetime.now(timezone.utc).isoformat()}
Scope: только последнее изменение sample results и непосредственно связанная регрессия.

1. VERDICT — PASS
   BLOCKER 0 / MAJOR 0 / MINOR 0. Дефектов нового автозаполнения не обнаружено.

2. CURRENT ARTIFACT VERIFIED — YES
   Exact path получен из implementation REPORT/manifest, не предположен:
   {artifact}
   SHA-256 {sha(artifact)}.
   Все 13 current product files совпали с final implementation tested BEFORE/AFTER/CURRENT,
   текущими повторными тестами и final current review hashes.
   Изменены только word.py, gui.py, STAGE5_DEV.md; добавлены result_samples.py и целевые тесты.
   Принятые core.py, inputs.py, Stage1–4, реестры и прежние suites не изменены.
   Runtime diff независимо связан с accepted before SHA: Word renderer и одна строка GUI.
   Guide: before/after SHA и полный current текст проверены; прежняя текстовая копия отсутствует.

3. HISTORICAL TEXT ASSESSMENT — PASS
   Независимое python-docx извлечение из 10 фактических исторических актов:
   56 физических строк, 56/56 «ДНО» в графе 7 и «годен» в графе 8.
   Регистр и литеральный текст сохранены; пустых/других вариантов=0.
   Полные source/mark/table/row/column records совпали с implementation extraction.
   Исторические источники не изменялись и не доказывают текущий контроль.
   Evidence: historical_rows_independent.json; historical_sources_independent.json.

4. NEW WORD RESULTS — PASS
   Ровно четыре строки/марки: Бп2-1, Пс1-9, Пс1-10, Пс2-16.
   В каждой физическая ячейка 7=ДНО, 8=годен. Смещения/лишних строк нет.
   Заголовки граф и первая шестёрка ячеек каждой строки byte-identical старому TEST.
   Package diff: изменён только word/document.xml. После возврата старого заголовка
   и ячеек 7/8 весь document.xml идентичен принятому старому TEST.
   Реквизиты, сварщики, оборудование, подпись, квалификация, настройки таблиц/страниц
   сохранены. Evidence: exact_document_columns.json; visual_validation_independent.json.

5. TEST-ONLY SAFETY — PASS
   Заметные жирные обозначения в заголовке:
   ТЕСТ / ОБРАЗЕЦ ЗАПОЛНЕНИЯ
   РЕЗУЛЬТАТЫ КОНТРОЛЯ НЕ ПОДТВЕРЖДЕНЫ
   Sample возвращает mode=SAMPLE_ONLY, confirmed=False. Официальные режимы STOP.
   Unsupported confirmation fields не превращают TEST в подтверждённый результат.
   Отсутствующий/изменённый источник STOP; изолированные blank/ambiguous/table mapping/
   row count/corrupt source также STOP. Drift после чтения блокирует запись файла.
   Будущий официальный режим требует фактических результатов, ответственного исполнителя,
   времени подтверждения и основания контроля по каждой марке; общая галочка количества
   этого не заменяет. Контракт не изменялся.

6. NATIVE WORD VALIDATION — PASS
   Microsoft Word {native['version']} / build {native['build']}; нормальное read-only
   открытие exact-файла, OpenAndRepair=False API default; закрыт без сохранения.
   Обе страницы native render просмотрены целиком в исходном разрешении 2480px.
   2 альбомные страницы / 1 секция / 2 таблицы / 9 строк первой таблицы / 0 hyperlinks.
   Четыре марки и обе последние графы читаются; предупреждения заметны.
   Обрезки, перекрытий, повреждённой таблицы, repair/recovery и лишних пустых страниц нет.
   Страница 2 сохраняет исполнителя, пустую линию подписи и квалификацию.
   Уже открытый пользовательский Word не закрывался и не сохранялся; его процесс сохранён.

7. CALCULATION REGRESSION — PASS
   Новый read-only preview текущего продукта: паспорт 9/2026, дата 01.10.2026.
   Бп2-1 Q8: 1280×5; 1024×8.4.
   Пс1-9 Q9: 3600×4.
   Пс1-10 Q15: 6000×4.
   Пс2-16 Q1: 480×3.
   Всего 4 отдельные марки, 33 изделия; все участки 5 МГц / 70°.
   Расчётные records совпали с принятыми; DOCX не создавался.

8. S5-R1 REGRESSION — PASS
   Четыре выбранных случая на новых isolated fixtures / actual current product: 4/4 PASS.
   NO_UZK_REQUIRED: 0 DOCX, 0 output directories, 0 writer calls, 0 source mutations.
   Два посторонних конфликта не блокируют; необходимый registry conflict вызывает STOP.
   GUI не разрешает генерацию при пустом пересечении.
   Evidence: A_empty_preview.json; B_no_output.json; s5r1_targeted.json/log.

9. INPUT / ARTIFACT INTEGRITY — PASS
   Все {len(protected)} защищённых файлов до/после review совпали по SHA.
   DEV ВЗМК/УЗК/Профили, проектный Справочник марок, ВОМ, первый шаблон,
   все 10 старых актов и старый TEST не изменены.
   Old TEST: {old['path']}
   Old SHA: {old['sha256']}.
   New SHA: {sha(artifact)}.
   Product не изменялся; production ВЗМК.xlsx / УЗК.xlsm не читались/не хешировались.
   Реестр не списывался, паспорт не закрывался. Mutations только в isolated fixtures.
   Implementation evidence: 281/281 SHA/size/exact file-set PASS, manifest-last PASS.
   Повторные targeted tests 16/16; новые independent checks 16/16; selected S5-R1 4/4.
   Новый REVIEW_MANIFEST.json записывается последним; далее только read-only verification
   всех его SHA/size/exact file set/last-write time и protected current files.

10. FINDINGS
   BLOCKER 0 / MAJOR 0 / MINOR 0. Findings: NONE.

11. NOT_RUN
   Фактический УЗК и подтверждение исполнителем по каждой марке; официальный выпуск;
   новая генерация акта/пересоздание exact TEST; полный Stage5 review/regression;
   Stage1–4 workflows; ручной USER E2E после изменения; production операции;
   списание реестра/закрытие паспорта; Stage6/deploy.

12. REQUIRED FIX — NONE

13. READY FOR USER E2E — YES
   Готовность относится к TEST/SAMPLE_ONLY. Ручной USER E2E ещё не выполнен.
   После review остановлено; код не исправлялся, новый акт не создавался.

CURRENT PRODUCT SHA-256 — implementation final tested BEFORE = AFTER = CURRENT = review final:
{product_lines}

INDEPENDENT SOURCE SHA / EXACT COUNTS:
{source_lines}

INDEPENDENT TARGETED CHECKS:
{negative_lines}

Implementation manifest SHA: {sha(IMPL/'REVIEW_MANIFEST.json')}
Hash tables: input_artifact_hashes_before_after.json; protected_before.json; protected_after.json.
Preparation-only resolved issues retained in PREPARATION_NOTES.txt and native_word preparation JSON.
'''
(OUT/'REPORT.txt').write_text(report,encoding='utf-8')

# Manifest is the final evidence write. Everything below it is read-only.
files={p.relative_to(OUT).as_posix():{'sha256':sha(p),'bytes':p.stat().st_size}
       for p in sorted(OUT.rglob('*')) if p.is_file() and p!=MANIFEST}
manifest={'created_utc':datetime.now(timezone.utc).isoformat(),'review':'Stage5 sample results targeted independent review',
          'status':'PASS','findings':findings,'required_fix':'NONE','ready_for_user_e2e':'YES',
          'mode':'SAMPLE_ONLY','results_confirmed':False,'artifact':str(artifact),'artifact_sha256':sha(artifact),
          'old_test_docx':old,'product_before':product_before,'product_current':product,
          'protected_count':len(protected),'protected_unchanged':True,'implementation_evidence_verified':281,
          'implementation_manifest_sha256':sha(IMPL/'REVIEW_MANIFEST.json'),
          'historical_sources':10,'historical_rows':56,'historical_pair':['ДНО','годен'],
          'targeted_tests':16,'independent_checks':16,'selected_S5R1_tests':4,'all_tests_passed':True,
          'native_word_validation':visual,'not_run':not_run,'evidence_files':files,'manifest_written_last':True}
save('REVIEW_MANIFEST.json',manifest)
listed=read(MANIFEST)['evidence_files']
actual={p.relative_to(OUT).as_posix():p for p in OUT.rglob('*') if p.is_file() and p!=MANIFEST}
assert set(actual)==set(listed)
for rel,e in listed.items():assert sha(actual[rel])==e['sha256'] and actual[rel].stat().st_size==e['bytes']
assert MANIFEST.stat().st_mtime_ns>=max(p.stat().st_mtime_ns for p in actual.values())
assert {p:sha(p) for p in protected}==protected
print(json.dumps({'status':'PASS','evidence_files_verified':len(files),'manifest_written_last_verified':True,
                  'manifest_sha256':sha(MANIFEST),'protected_files_unchanged':len(protected),
                  'targeted':'16/16 PASS','independent':'16/16 PASS','S5R1':'4/4 PASS',
                  'report':str(OUT/'REPORT.txt')},ensure_ascii=False,indent=2))
