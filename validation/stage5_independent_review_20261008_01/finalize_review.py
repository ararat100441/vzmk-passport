import hashlib,json
from datetime import datetime,timezone
from pathlib import Path
RUN=Path(__file__).resolve().parent;ROOT=RUN.parents[1]
OLD=ROOT/'validation/stage5_implementation_20261008_105100'
def read(p):return json.loads(Path(p).read_text(encoding='utf-8-sig'))
def sha(p):return hashlib.sha256(Path(p).read_bytes()).hexdigest()
def write(n,v):(RUN/n).write_text(json.dumps(v,ensure_ascii=False,indent=2,default=str),encoding='utf8')
initial=read(RUN/'review_initial_hashes.json')
assert all(Path(p).name.casefold() not in ('взмк.xlsx','узк.xlsm') for p in initial)
after={p:sha(p) for p in initial};assert initial==after
write('review_final_hashes.json',after)
m=read(OLD/'EVIDENCE_MANIFEST.json');integrity=read(RUN/'integrity.json')
assert all(sha(OLD/n)==v['sha256'] for n,v in m['evidence_files'].items())
product_before=read(OLD/'product_before.json')['accepted_product']
assert all(sha(p)==h for p,h in product_before.items())
write('accepted_stage1_4_product_comparison.json',{p:dict(before=h,after=sha(p),match=True) for p,h in product_before.items()})
assert all(p['match'] for p in read(RUN/'accepted_stage4_product_comparison.json'))
native=read(RUN/'native_alerts_enabled.json');first_native=read(RUN/'native_exact.json')
assert native['status']=='PASS_NORMAL_READONLY_OPEN' and native['unchanged'] and native['pages']==2
assert native['alerts_enabled'] and native['explicit_OpenAndRepair'] is False
assert all(sha(RUN/f'native_exact_{p}.png')==sha(RUN/f'native_alerts_enabled_{p}.png') for p in (1,2))
write('visual_review.json',dict(document=m['final_output'],native_word_read_only=True,normal_open=True,OpenAndRepair=False,alerts_enabled=True,repair_warning_or_prompt='none: normal open completed with warnings enabled',pages=2,all_pages_visually_inspected=True,landscape_A4=True,one_section=True,two_tables=True,physical_mark_rows=4,page1='All four marks, five geometry segments, frequency and angle aligned, full table boundaries readable; no clipping or overlap',page2='Original operator, empty signature line, qualification text fully readable; no extra blank page',both_page_hashes_equal_to_images_visually_inspected=True,inspected_pngs=[str(RUN/f'native_exact_{p}.png') for p in (1,2)],source_docx_unchanged=True,PDF_created=False))
suite=read(RUN/'rerun_35.json');probes=read(RUN/'supplemental_probes.json');extras=read(RUN/'extra_probes_attempt0.json')[:3]
assert suite['tests_run']==35 and suite['successful']
assert sum(p['status']=='PASS' for p in probes)==18 and len(probes)==19
assert all(e['status']=='PASS' for e in extras)
repro=read(RUN/'S5-R1_reproduction.json');assert repro['reproduced'] and repro['fixture_unchanged']
findings=[dict(ID='S5-R1',severity='MAJOR',location='uzk_stage5/core.py:96-100; uzk_stage5/inputs.py:192-199',problem='Preview validates unrelated mark-catalog rows before establishing that the shipment requires no UZK.',reproduction='Retained isolated no_uzk_True: valid VOM, 25 valid registry rows all belonging to OTHER_VALID_PROJECT, no selected project/mark matches, two conflicting UNRELATED_MARK catalog rows. Full preview raises MARK_PROFILE_CONFLICT at row 198 instead of returning zero marks.',evidence=['supplemental_probes.json','S5-R1_reproduction.json','scenarios/no_uzk_True'],impact='Valid no-UZK shipments are blocked by unrelated catalog conflicts. No wrong act or source mutation observed.',minimal_fix='After mandatory valid VOM and registry read, determine the project+mark intersection. For an empty intersection return a sealed empty preview before unrelated profile mapping/rule validation; generate must return None without creating a directory. Missing/unreadable registry must still STOP.',targeted_validation='Exercise full preview->generate with zero UZK matches and unrelated conflicting/unresolved catalog rows. Assert zero marks, no DOCX or directory, unchanged registry. Keep missing/unreadable-registry STOP and relevant-profile-conflict STOP tests. Existing test_29 only replaces data[marks] with [] and bypasses preview.')]
write('FINDINGS.json',findings)
inp=read(RUN/'independent_inputs.json');trace=read(RUN/'independent_four_mark_trace.json');doc=read(RUN/'docx_structure.json')
products='\n'.join(f"- {p['path']}\n  SHA256 {p['sha256']}" for p in read(OLD/'changed_files.json')['new_product_files'])
inputs='\n'.join(f'- {p}\n  SHA256 {h}' for p,h in m['inputs'].items() if Path(p).suffix in ('.xlsx','.xlsm','.xls','.docx'))
rows='\n'.join(f"{r['mark']} | ВОМ {r['vom_quantity']} (стр. {r['vom_rows'][0]['row']}) | остаток {r['remaining']} (реестр стр. {r['registry_row']}) | Q={r['Q']} | {r['raw_profile']} (каталог стр. {r['catalog_row']}) | "+'; '.join(f"{s['L']}×{s['S']}" for s in r['final']) for r in trace)
source_rows='\n'.join(f"- {r['mark']}: {r['normalized_key']} / {r['standard']}; {r['rule']}; геометрия {r['source_geometry']}" for r in trace)
report=f'''ВЗМК — STAGE 5 INDEPENDENT REVIEW
Дата: 08.10.2026, Europe/Samara

1. VERDICT — FAIL
BLOCKER 0 / MAJOR 1 / MINOR 0.
Проверяемый DOCX и основной расчёт проходят; полный контракт Stage 5 не выполнен из-за S5-R1.

2. ACTUAL STATE
Exact TEST DOCX:
{m['final_output']['path']}
SHA256 {m['final_output']['sha256']}
Файл установлен из REPORT.md и EVIDENCE_MANIFEST.json, не угадан и не заменён новым.
SHA совпал до проверки и после двух native Word открытий без сохранения.

Exact Stage 5 product set: 10 файлов (8 runtime/launcher/dependencies, 1 test, 1 instruction):
{products}
Все 10 SHA совпали с changed_files.json. Tested set из 9 файлов включает тест и не включает STAGE5_DEV.md; его BEFORE=AFTER=CURRENT, инструкция отдельно связана SHA.

Exact source fingerprints:
{inputs}
Checkpoint, settings, три Stage 4 sidecar JSON и все связи входов дополнительно проверены в independent_inputs.json / integrity.json.
Производственные ВЗМК.xlsx / УЗК.xlsm не открывались и не хешировались.

3. SOURCE / INPUT ASSESSMENT
ВОМ XLS: 104 видимые строки, 397 скрытых; масса 15822.14 кг. Независимое чтение native Excel по фактическим B/D/I и SpecialCells(12), а не доверие UsedRange/расчётному JSON. Сумма COM double сохранена отдельно: 15822.13999999999951; округление до 0.01 даёт ожидаемое значение.
Входной AutoFilter $A$2:$K$507. Отдельные XLS/XLSX fixtures с AutoFilter и manual-hide прошли повторный исходный test_04. Скрытые количества и масса не суммируются. Повторные марки агрегируются безопасно.
Реестр УЗК: 25 записей A2:C26, AutoFilter A1:C21; строки 22–26 читаются. C пусто -> 1, валидное положительное целое -> C, Q=min(ВОМ,остаток). Identity проект+марка; одинаковый дубликат не увеличивает остаток, конфликт STOP. Проверены восемь недопустимых C через настоящий workbook reader.
Справочник проекта: фактические raw строки 40/78/79/98 независимо прочитаны; accepted Stage 4 workbook+sidecars проверены.

4. FOUR-MARK RECONCILIATION
{rows}
Одно изделие: Бп2-1 160×5 и 128×8.4; Пс1-9 и Пс1-10 400×4; Пс2-16 480×3.
Всего 4 марки / 33 единицы / 5 сегментов. Все сегменты 5 МГц / 70°. Толщина S не умножается. Разные марки не объединяются.
Stage 4 lookup / rules:
{source_rows}
Все четыре lookup FOUND. Stage 5 вызывает принятые Catalog.lookup/calculate_uzk; собственной геометрической формулы нет. Raw profile сохранён в business preview/output JSON; шаблон DOCX не содержит отдельной графы профиля, и её не добавляли.

5. WORD TEMPLATE / FIELD ASSESSMENT — PASS
Детерминированно выбран ровно первый паспорт 1_2026 - УЗК.docx, SHA приведён выше. A4 landscape, одна секция, две таблицы.
Строка паспорта найдена динамически: {inp['passport']['row']}, без hardcoded 400.
Проект {inp['passport']['project']}; номер {inp['passport']['number']}; дата N400 — 01.10.2026.
N содержит {inp['passport']['formula_N']}; прочитано принятое cached значение, формула не пересчитывалась.
E400: {inp['passport']['object']}
F400: {inp['passport']['cipher']}
Эти значения присутствуют в точном DOCX. Дата не заменена текущей/VIK/датой формирования.
Из пакета изменены только document.xml и его hyperlink relationships. Signature/qualification table и sectPr побайтно равны шаблону; постоянные company/equipment/operator/welder/qualification данные сохранены. Результаты дефектов/заключения пусты. Видимая пометка ТЕСТ / ЧЕРНОВИК присутствует.

6. WORD ROW / HYPERLINK ASSESSMENT — PASS
Ровно 4 физические строки: Бп2-1, Пс1-9, Пс1-10, Пс2-16; по 8 логических ячеек. Бп2-1 занимает одну строку с двумя согласованными строками geometry/frequency/angle.
Исторические марки и placeholders отсутствуют. Гиперссылки Word=0; hyperlink relationships=0; external PDF targets=0; hidden stale mark/field parts не обнаружены.

7. NATIVE WORD VISUAL VALIDATION — PASS
Точный файл открыт установленным Word read-only обычным Documents.Open, OpenAndRepair=False явно, предупреждения включены. Открытие завершилось без repair/recovery/corruption сообщения и без API восстановления.
Две страницы. Обе native PNG страницы реально просмотрены: стр. 1 все четыре марки и границы/сегменты читаемы; стр. 2 исходная строка подписи и квалификация читаемы. Landscape; нет clipping, overlapping cells, blank historical rows или лишней пустой страницы.
Снимки повторного warning-enabled открытия побайтно совпали с просмотренными PNG. DOCX не сохранён, SHA не изменён. PDF не создавался.

8. PREVIEW / GENERATION SAFETY — PASS EXCEPT S5-R1
Свежий read-only preview дал тот же seal, что final_run.json. Изменения после preview отдельных копий ВОМ / УЗК registry / project catalog / profile database / DEV passport registry / template вызывают INPUT_CHANGED до создания выхода.
Исключительное создание xb предотвращает overwrite/race. Existing file и настоящий Windows share=0 lock сохранены. Смоделированные denied-open и partial write удаляют только свой частичный файл. Изолированный успешный документ имеет номер 777/2099 и явный fixture object; реальный 9/2026 не генерировался повторно.
Tk gate проверен программно: preview + checkbox обязательны; смена выбора инвалидирует plan; busy запрещает generation. Open output callback проверен на точный путь, сам файл реально открыт native Word.
Чистый end-to-end no-UZK возвращает ноль марок и None без папки/документа; с посторонним конфликтом каталога обнаружен S5-R1. Missing/unreadable registry STOP, не превращается в no-UZK.

9. REGISTRY / SOURCE MUTATION SAFETY — PASS
Stage 5 input write/decrement/close путей не обнаружено. Excel DispatchEx, ReadOnly=True, AutomationSecurity=3, EnableEvents=False, UpdateLinks=0, AddToMru=False; внешние connections/links отсутствуют; Close(False)/Quit. Failure-open cleanup проверен с COM mock. openpyxl источники read-only, keep_links=False.
Native Word работал в отдельном процессе с нулём исходных документов, закрыт без save. Чужие Office сеансы не закрывались. Пустой automation-owned Word от неудачной инициализации review renderer закрыт только после проверки ownership/time, Documents.Count=0, UserControl=False, Visible=False.
Все {len(initial)} наблюдаемых исходных/продуктовых/evidence отпечатков REVIEW BEFORE=AFTER. Исходные 63 protected файла implementation BEFORE=AFTER=CURRENT. В runtime Python guard production reads=0, writes outside review=0. Все fixture mutations/output внутри этой новой папки.
Нормальный preview/generation Python network attempts=0; отдельный resolver guard остановил вызов до сети. Excel источники не имеют внешних подключений. Live web/source audits не выполнялись.

10. NEGATIVE TESTS
Сохранённые implementation 36/36 PASS подтверждены кодом тестов, ключевыми результатами и SHA evidence.
Независимо повторены 35/35 исходных тестов, 0 skips. Не запускался один test_32_33, который повторно генерирует реальный 9/2026; заменён свежим read-only preview, access audit и isolated generation.
Дополнительно 18/19 targeted probes PASS; один FAIL = S5-R1. Отдельно 3/3 завершённых extras: invalid registry quantities through workbook, real Windows locked existing output, blocked resolver. Отдельный повторяемый S5-R1 repro подтвердил 25 valid registry records, 0 selected project matches, ошибку и 0 DOCX/fixture mutations.
Покрыты missing project, ambiguous passport, missing template/registry, unreadable registry, conflicting registry/catalog, invalid quantity, missing/ambiguous/conflicting profile, unsupported rule, hidden rows, stale inputs, failed save, existing output/race, no-UZK.
Логи неудачных harness attempts сохранены. Первый temp writer был остановлен guard вне review; temp перенаправлен внутрь review, повтор 35/35 прошёл. Независимый End(xlUp) пропускал 14 скрытых хвостовых строк; исправлен только review reader на stored range + SpecialCells, подтверждены 397 hidden. Поздний избыточный повтор extra helper найден suspended и остановлен по точной process identity; он не засчитан. Успешные ранее завершённые extras сохранены в extra_probes_attempt0.json.

11. EVIDENCE INTEGRITY — PASS
57/57 исходных evidence SHA и размеров совпали; отсутствующих/лишних файлов нет. Манифест написан последним относительно всех 57 файлов. Итоговый DOCX 07:06:15 UTC, native proof 07:06:20, tested BEFORE 07:06:29, AFTER/test summary 07:06:38, manifest 07:10:49 — порядок согласован. Product tested BEFORE=AFTER=CURRENT; output fingerprint chain соответствует текущему product, свежий seal равен сохранённому.
Старые QA iterations/failed attempts сохранены и не выданы за final_output_v3.
Старый manifest SHA {integrity['manifest_sha256']} не изменён review.
Финальный новый REVIEW_MANIFEST.json связывает этот отчёт, findings, исполнявшиеся harness scripts, все fixtures/outputs и финальные source hashes; себя не хеширует.

12. REGRESSION IMPACT — PASS WITHIN REQUESTED BOUNDARY
30 accepted Stage 1–4 product fingerprints совпали с implementation BEFORE и CURRENT. 10 текущих Stage 4 product hashes отдельно совпали с accepted targeted-review manifest 07.10.2026. main/stage2/stage3/stage4 imports + DEV pins PASS. Stage 5 — отдельный модуль; чужие workflow не выполнялись.
index_status/codebase_context tools не доступны в этой сессии; выполнено ограниченное чтение фактического source set и прямых зависимостей, без веб-поиска.

13. FINDINGS
S5-R1 — MAJOR.
Location: uzk_stage5/core.py:96–100 (безусловный mark_profiles до calculate_marks); inputs.py:192–199 (проверка конфликтов всех строк).
Problem: отсутствие марок УЗК не предотвращает проверку посторонних конфликтов каталога.
Reproduction: сохранённая scenarios/no_uzk_True. Все 25 корректных registry records относятся к OTHER_VALID_PROJECT; selected PROJECT+MARK intersection пуст. Две UNRELATED_MARK строки с разными профилями. Full preview -> MARK_PROFILE_CONFLICT: UNRELATED_MARK, строка 198.
Evidence: supplemental_probes.json; S5-R1_reproduction.json; reproduce_no_uzk.py; исходные fixture bytes.
Impact: корректный no-UZK shipment ошибочно блокируется; неверного акта/списания не происходит.
Minimal fix: после обязательного чтения валидных VOM/registry определить intersection. При пустом intersection вернуть sealed empty preview до проверки нерелевантных profile mappings/rules. Missing/unreadable registry должен оставаться STOP. Для непустого intersection сохранить fail-closed проверки выбранных марок.
Targeted validation: полный preview->generate при пустом intersection + unrelated conflict/ПРОВЕРИТЬ; 0 marks, no directory/DOCX, unchanged registry. Проверить missing/unreadable-registry STOP и selected conflict STOP. Исходный test_29 искусственно задаёт data[marks]=[] и не проверяет этот путь.
BLOCKER 0 / MAJOR 1 / MINOR 0.

14. NOT_RUN
- Полный Stage 1–4 regression и workflows, генерация passport/VIK.
- Пользовательский GUI E2E (программные Tk gate проверки выполнены).
- Повторная генерация реального DEV документа 9/2026 (намеренно исключена).
- Live ultrasonic inspection, результаты, sign-off.
- Production registry reads/hashes/writes; decrement/delete/close.
- PDF/EXE/deploy/Stage 6; live web/source audit.
- Поздний избыточный suspended helper repeat не завершён/не засчитан; исходные завершённые extras и отдельный полный S5-R1 repro имеются.

15. REQUIRED FIX
Устранить S5-R1 и выполнить названный targeted end-to-end no-UZK validation. Implementation не изменена в этом review.

16. READY FOR USER E2E — NO
Review остановлен. Stage 6 не начинался.
'''
(RUN/'REPORT.txt').write_text(report,encoding='utf8')
write('review_summary.json',dict(verdict='FAIL',blocker=0,major=1,minor=0,required_fix='S5-R1',ready_for_user_e2e=False,review_observed_hashes_unchanged=len(initial),original_protected_unchanged=63,original_evidence_hashes_verified=57,original_tests_rerun=dict(pass_count=35,total=35),supplemental=dict(pass_count=18,total=19),completed_extras=dict(pass_count=3,total=3),exact_docx=m['final_output'],actual_generated_document_changed=False,source_writes=0,production_registry_reads=0,full_prior_regression='NOT_RUN'))
evidence={str(p.relative_to(RUN)):dict(sha256=sha(p),bytes=p.stat().st_size) for p in sorted(RUN.rglob('*')) if p.is_file() and p.name!='REVIEW_MANIFEST.json'}
write('REVIEW_MANIFEST.json',dict(created_utc=datetime.now(timezone.utc).isoformat(),verdict='FAIL',findings=findings,exact_test_docx=m['final_output'],review_source_before=initial,review_source_after=after,source_changes=[],original_implementation_manifest_sha256=sha(OLD/'EVIDENCE_MANIFEST.json'),evidence_files=evidence,manifest_written_last=True))
print(json.dumps(dict(report=str(RUN/'REPORT.txt'),verdict='FAIL',blocker=0,major=1,minor=0,review_hashes_unchanged=len(initial),review_evidence_files=len(evidence),ready_for_user_e2e=False),ensure_ascii=False,indent=2))
