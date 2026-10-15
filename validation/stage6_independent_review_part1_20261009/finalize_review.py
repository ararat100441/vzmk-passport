"""Finalize independent review. REVIEW_MANIFEST.json is written last."""
from review import *
from datetime import timezone,timedelta
before=load(E/'hash_before.json');impl=load(I/'EVIDENCE_MANIFEST.json')
tests=load(E/'independent_tests.json');reconciliation=load(E/'reconciliation.json')
binding=load(E/'actual_output_product_binding.json');historical=load(E/'accepted_historical_hashes.json')
after={p:sha(p) for p in before['all']}
assert after==before['all'] and not before['mismatches']
assert tests['summary']=={'PASS':53,'FAIL':0} and binding['seal_exact_match']
assert not tests['audit']['production_attempts'] and not tests['audit']['network_attempts']
assert load(E/'locked_output_gui_reproduction.json')['safe_rollback']
guarded=dict(after)
for batch in historical.values():
    for path,digest in batch['current'].items():
        assert sha(path)==digest
        guarded[path]=digest
historical_docs={p:h for p,h in historical['audit']['current'].items()
                 if Path(p).is_relative_to(ROOT.parent/'Гараж казахстан/Паспорт')}
code_set=[ROOT/'akz_stage6/__init__.py',ROOT/'akz_stage6/core.py',ROOT/'akz_stage6/gui.py',ROOT/'stage6.py',
          ROOT/'Запустить Этап 6 DEV.cmd',ROOT/'STAGE6_DEV.md',ROOT/'tests/test_stage6.py']
support=list(ROOT.glob('validation/stage6_*.py'))
reader_set=[ROOT/'passport_app/registry.py',ROOT/'passport_app/excel.py',ROOT/'passport_app/tables.py',ROOT/'passport_app/config.py']
fixture=Path(load(I/'test_assignments.json')['fixture'])
save('product_set.json',{'product':{str(p):sha(p) for p in code_set},'accepted_read_dependencies':{str(p):sha(p) for p in reader_set},
     'implementation_support_scripts':{str(p):sha(p) for p in support},
     'final_fixture_set':{str(p):sha(p) for p in fixture.parent.iterdir() if p.is_file()},
     'other_implementation_fixtures_bound_in_original_manifest':[p for p in impl['files'] if 'Stage6_Fixtures' in p]})
findings=[
 {'id':'S6-P1-R1','severity':'MINOR','location':str(ROOT/'akz_stage6/core.py')+':406; '+str(Path(impl['canonical_output'])/'ТЕСТ Приложение к АКЗ №0.xlsx')+'!A7:G20',
  'reproduction':'Open the exact application READ-ONLY in native Excel and print preview. Rows 4:6 contain the three selected marks; rows 7:20 are blank and visible; totals remain in row 21. Native snapshots and print screenshot confirm 14 blank table rows.',
  'impact':'Violates the explicit Part 1 requirement that the working table contain no extra blank rows. Mass, area and quantity totals are correct; no data loss.',
  'minimal_fix':'On generated copies hide unused body rows between the last selected row and row 21, or otherwise compact the table without altering the template source or removing selected rows.',
  'targeted_validation':'Three selected rows: exactly three visible data rows followed by totals in native print preview. Seventeen selected rows: all retained. Eighteen rows: explicit STOP. Verify formulas/totals and original template hashes.'},
 {'id':'S6-P1-R2','severity':'MINOR','location':str(ROOT/'akz_stage6/core.py')+':133; '+str(ROOT/'akz_stage6/gui.py')+':83',
  'reproduction':'In the isolated pending pair hold the valid second XLSX with CreateFile GENERIC_READ, sharing mode 0, before native_open. Run the real Launcher.generate handler. GUI reports that the format or extension is invalid although the copied file matches the valid pinned application template SHA.',
  'impact':'Safe STOP and full rollback work, but the user receives a misleading corruption/format message for an output-file lock, contrary to the required clear lock error.',
  'minimal_fix':'Wrap native open/save failures with a clear file/operation error identifying a busy or inaccessible output and a close/retry instruction; preserve the original diagnostic internally and retain pair rollback.',
  'targeted_validation':'Lock the valid second pending workbook: clear busy/access error, no published first workbook, no pending pair. Release lock: complete pair succeeds. Existing outputs remain unchanged.'}
]
save('findings.json',{'verdict':'FAIL','counts':{'BLOCKER':0,'MAJOR':0,'MINOR':2},'findings':findings})
save('hash_after.json',{'all':after,'changes':[],'product_BEFORE_AFTER_CURRENT_equal':True,
      'guarded_union':guarded,'guarded_count':len(guarded),'historical_documents_count':len(historical_docs),
      'historical_documents_unchanged':True,'outputs':before['outputs']})
save('HARNESS_RECOVERY.json',{'review_harness_only':True,
    'initial_failure':'The first synthetic boundary fixture serialization dropped namespaces referenced by mc:Ignorable; Excel refused that reviewer-created fixture.',
    'repair':'Only review harness XML serialization was switched to namespace-preserving lxml. Continued from the first incomplete 17/18 checkpoint, retaining prior completed checks.',
    'product_or_real_source_changes':0,'negative_output_scope':str(E/'sandbox'),
    'lock_check_scope_correction':'The initial native lock test established a safe error/rollback only. Targeted real GUI reproduction separately established that error clarity FAILs (S6-P1-R2). No clarity PASS is claimed.'})
stamp=datetime.now(timezone(timedelta(hours=4))).isoformat()
out=Path(impl['canonical_output'])
sources=load(I/'source_integrity_current.json')['BEFORE']
full_product='\n'.join(f'{p}\n  SHA-256 {sha(p)}' for p in code_set)
full_support='\n'.join(str(p) for p in support)
report=f'''1. VERDICT — FAIL

Независимая проверка Stage 6 Part 1, {stamp}, Europe/Samara.
BLOCKER 0; MAJOR 0; MINOR 2. Проверка завершена, исправления не выполнялись.

2. ACTUAL STATE

Рабочая папка: {ROOT}
Фактический исполняемый продукт, launcher, tests, documentation:
{full_product}

Core содержит current_passport/catalog/visible_vom, matching/preview и генерацию Excel.
Принятые read dependencies: passport_app/registry.py, excel.py, tables.py, config.py.
Stage 6 support / evidence drivers:
{full_support}
Полный набор и все хеши: product_set.json. Итоговая DEV fixture: {fixture}

Точные проверяемые XLSX получены из implementation manifest, не из предположения об имени:
ACT: {out/'ТЕСТ АКЗ №0.xlsx'}
SHA-256 1a52736427e1529420723ff85f3bfd92340865ece895173ffaaa0dadb48ba06f
APPLICATION: {out/'ТЕСТ Приложение к АКЗ №0.xlsx'}
SHA-256 4cd657f84311512255e6a6e39b492c227f18df3ab8b5ab03eb34400f165d9c6d
Оба SHA проверены до тестов, после native Excel и в финале; полное совпадение.

3. TEMPLATE PAIR ASSESSMENT — PASS

ACT: {ROOT.parent/'Гараж казахстан/Паспорт/Акт № 4-2026.xlsx'}
SHA-256 7f6420ac50149805b46454852d62ad68fa9625c53b7c9d375460769b6ae0f8ac
APPLICATION: {ROOT.parent/'Гараж казахстан/Паспорт/1_2026 - Приложение к АКЗ.xlsx'}
SHA-256 704d92eb0c5ea87c3951f6643488aa9100cd18f99e177d21ddc85f8843938d6f
Подтверждены внутренние 1/2026, 08.09.2026, ACT 4!G8=E4 / I8=H4 и согласованный A1 приложения.
Выбор задан явными путями и pins; fuzzy matching отсутствует. Число 4 в имени не участвует в новой нумерации.
Шаблоны других объектов не используются. Исходные шаблоны неизменны.

4. DEV TEST ASSIGNMENTS — PASS

Реальный справочник: {ROOT.parent/'Гараж казахстан/Справочник марок.xlsx'}
195 марок, все назначения blank. Он не изменён.
Отдельная fixture: {fixture}
SHA-256 7bd30ae4dae52fe9ad43e1d86debef825eca77be1fc14f0feb92245b03483b03
Бп2-5 / Бп2-18 / Бп2-19: Марки!C41 / C47 / C48, числовые n=0.
192 остальных значения остаются None. Все исходные марки и профили A:B совпадают построчно.
Blank не приравнивается к 0; 1/2/3 исключаются; неподдерживаемое 4 вызывает STOP без fallback.

5. VOM / VISIBLE ROW ASSESSMENT — PASS

{reconciliation['vom']['path']}
SHA-256 aae5e5b8c0246c8fe97b3dec9cbcc039bc3e3cedb9d245b474b2f871b90596ae
Лист !01_ВЗМК_A_Ведомость отправочны; AutoFilter $A$2:$K$507, активен фильтр количества.
Независимый native read-only snapshot: 501 строка данных = 104 видимые + 397 скрытых.
Марка B, имя C, количество D, масса H:I, площадь J:K. Заголовки проверены непосредственно в книге.
Все 104 номера строк и исходные значения совпали с reader; totals полного UsedRange не используются.
XLS и XLSX isolated visibility fixtures подтвердили исключение фильтрованных и вручную скрытых строк.
Повторные марки сохраняются отдельными строками и в исходном порядке.

6. THREE-MARK RECONCILIATION — PASS

MARK | VOM ROW | QUANTITY | UNIT MASS kg | TOTAL MASS kg | UNIT AREA m2 | TOTAL AREA m2
Бп2-5 | 40 | 1 | 37.2 | 37.2 | 1.412 | 1.412
Бп2-18 | 53 | 1 | 31.32 | 31.32 | 1.184 | 1.184
Бп2-19 | 54 | 1 | 36.74 | 36.74 | 1.393 | 1.393
ИТОГО: 3 изделия; 105.26 кг; 3.989 м2.
Все raw марки точны, наименование каждой — Балка; все три видимы и имеют назначение numeric 0.
Каждая пара unit * quantity = total независимо проверена Decimal из native VOM values.
Альтернативные приложения, глобальная ВОМ, Stage 4 geometry и исторические числа не использованы.
Точность площади в XLSX сохранена (0.000); Excel double 105.26000000000002 не означает бизнес-округления.

7. ACT FIELD ASSESSMENT — PASS

Принятый read_registry/select_project: первый 1/2026, row 353; текущий 9/2026, row 400.
Дата из DEV листа Паспорт качества, O1=Дата Акз, O400 cached=05.10.2026.
O400 содержит формулу =Таблица1[[#This Row],[Дата окончания ]]-1; читается именно именованное поле АКЗ.
Папка 9 от 06.10.2026, компьютерная дата, исторический акт, ВИК и УЗК не определяют дату результата.
ACT E4=9/2026 от; H4=05.10.2026; G8/I8 согласованы с приложением.
Договор C400: 06/26 от 03.06.2026; заказчик D400: ООО "СТК ВК".
Объект E400: Архитектурно-строительные решения здания парковки подвижного состава.
КМД F400: WAA0007144-AC2-КМД.
ЛКМ Pirocor Protect 057, RAL 7004, один слой и требуемые 80 мкм сохранены из pin-шаблона.

8. APPLICATION ASSESSMENT — FAIL, S6-P1-R1

A:G содержат марку, наименование, количество, массу одной/всех, площадь одной/всех.
Строки 4:6 — только три текущие марки в точном порядке. Исторические марки очищены.
A1 содержит 9/2026 и 05.10.2026; ACT-ID обеих книг e9fc12e4433feafd.
C21/E21/G21 = SUM(C4:C20)/SUM(E4:E20)/SUM(G4:G20); stored и native totals верны.
Однако A7:G20 — 14 пустых видимых строк внутри рабочей таблицы, они печатаются до итогов.
Это прямо нарушает требование об отсутствии лишних пустых строк. Данные и итоги не теряются.

9. 17-ROW LIMIT ASSESSMENT — PASS

Причина: исходная форма выделяет 17 строк A4:G20 и итоги в row 21; matching явно ограничивает len<=17.
Независимая synthetic 17-row fixture: preview сохраняет все 17 и их порядок; native generation сохранила все 17.
Контрольные итоги synthetic: 17 шт., 20.91 кг, 7.752 м2. Это не данные реального паспорта.
18 строк: TEMPLATE_CAPACITY_STOP до генерации; неполный акт не создан.
Ограничение явно указано STAGE6_DEV.md:14-15 и в сообщении STOP интерфейса.

10. HISTORICAL MEASUREMENTS / CLIMATE SAFETY — PASS

Все 24 фактических поля проверены stored formula/cache и после native CalculateFullRebuild.
Исторические 92 мкм, мокрая толщина, климат, партии, даты/время нанесения и результаты не переносятся.
В них НЕТ ДАННЫХ / ТЕСТ или ?; остаются только формулы G8=E4 и I8=H4.
Нет формул, превращающих placeholders в фиктивные толщины/климат; RANDBETWEEN отсутствует.
Оба файла заметно помечены ТЕСТ / ЧЕРНОВИК; акт сообщает, что выдача запрещена.
ClimateDay helpers проверены для одной записи на день, конфликта и Ts>=Td+3 на synthetic unit tests.
В exact output все измерения MISSING. Формула Td не выбирается. Официальный режим вызывает STOP.
Реальный climate engine не требуется и не объявляется реализованным.

11. NATIVE EXCEL VISUAL VALIDATION — PASS с дефектом таблицы из пункта 8

Microsoft Excel 16.0; обе exact книги открыты обычным CorruptLoad=0, UpdateLinks=0, ReadOnly=True.
Макросы и события выключены, книги не сохранялись; SHA после закрытия совпали.
Repair/recovery для exact output не наблюдались; по одному ожидаемому листу.
Номер, дата, реквизиты, марки, числа и placeholders читаемы; обрезание данных печатными границами не обнаружено.
Native print preview: ACT 1 из 1, B2:K50; APPLICATION 1 из 1, A1:G21.
Печатный вид приложения подтверждает лишние пустые строки, а не отсутствие этого дефекта.
Evidence: native_act_print.png/txt, native_application_screen.png, native_application_print.png/txt,
native_outputs_on_open.json и native_outputs_recalculated.json. Физической печати не было.

12. PREVIEW / PAIR GENERATION SAFETY — safe generation PASS; error clarity FAIL, S6-P1-R2

Preview не изменяет источники. Реальные Tk widgets: генерация до preview запрещена;
изменение project/passport/VOM/catalog/synthetic сбрасывает checked и отключает кнопку.
Изменение каждого из registry/VOM/catalog/ACT/APPLICATION после preview вызывает STOP.
Изменение источника во время генерации также приводит к STOP до публикации.
Существующий output не перезаписывается. Ошибка второй книги и ошибка saved validation откатывают pending.
Пара публикуется одним rename только после сохранения, повторного read-only открытия и проверки обеих книг.
Actual OS lock второй книги: неполной публикации и утечки pending нет.
Но GUI показывает недостоверное сообщение о неверном формате/расширении исправного заблокированного XLSX.

13. NEGATIVE TESTS

53 адресные runtime assertions PASS; неизменённый product unit suite 16/16 PASS.
A blank !=0 PASS. B 1/2/3 excluded PASS. C conflicting duplicate STOP PASS.
D отсутствие exact марки при отсутствии выбранной группы: NO_AKZ_ZERO_ROWS STOP PASS;
невыбранные/неназначенные марки показываются как NO_ASSIGNMENT, без переименования.
E missing/invalid/zero/negative area STOP PASS, включая native preview fixture.
F inconsistent unit/total area STOP PASS, включая native preview fixture.
G missing/changed templates STOP PASS. H existing output preserved PASS.
I changed five inputs / altered payload / mid-generation change STOP PASS.
J 18 rows STOP PASS, 17 rows complete PASS. K no selected numeric zero STOP PASS.
Это не означает PASS двух требований оформления и понятной ошибки: их фактические FAIL записаны findings.json.
Все mutation tests — {E/'sandbox'}; реальные выходные книги для negative tests не создавались.
Одна ошибка reviewer-created XML fixture исправлена только в review harness; подробности HARNESS_RECOVERY.json.

14. SOURCE / PRODUCTION SAFETY — PASS

82 source baseline files реализации unchanged; original manifest 195/195 текущих fingerprints совпали до/после.
Stage 1–5 не менялись относительно accepted baseline перед Stage 6. Реальные DEV XLSX/XLSM,
Профили.xlsx, каталог, selected VOM, обе исходные формы и проверяемая output-пара неизменны.
Исторические документы выбранного проекта: {len(historical_docs)}/{len(historical_docs)} совпали с accepted audit hashes.
92/92 fingerprints финальной проверки шаблонов совпали.
Старый audit содержит четыре до-Stage6 изменения: catalog.py, catalog_gui.py, uzk_stage5/inputs.py и real catalog.
Все четыре точно объяснены accepted stage3_catalog_schema_20261008_162454 evidence и current SHA.
Ни одно из них не является изменением Stage 6 или этой review.
Production registries: ZERO reads/writes. Python audit попыток production/network: 0/0.
Вызовов сетевых инструментов не было; статически сетевых вызовов продукта нет; native source open отвергает external data.
ROOT/DEV/output/fixtures/project — реальные директории без junction/reparse перенаправления.
Изменения review ограничены новой папкой evidence. Stage 7 не выполнялся.

15. EVIDENCE INTEGRITY — PASS

Все пять записей tested_product BEFORE/AFTER/CURRENT точно совпали с исполняемыми файлами.
Launcher/docs и восемь support scripts связаны с исходным manifest и current SHA.
Полный current preview на original inputs повторно дал исходный seal:
{binding['current_product_seal']}
Он совпадает с implementation preview и PAIR ACT-ID; current product validate_saved_pair прошёл на exact output read-only.
Независимые snapshots, пересчёт, negative tests, Tk проверки и находки сохранены в этой папке.
195 baseline fingerprints и расширенный union {len(guarded)} guarded files неизменны.
Новый REVIEW_MANIFEST.json записывается последним после report/findings/hash_after и проверяется read-only.
Codebase index tools отсутствовали; применены адресные чтения фактических файлов. Непроведённые проверки не названы PASS.

16. FINDINGS

BLOCKER 0
MAJOR 0
MINOR 2

S6-P1-R1 | MINOR
LOCATION: {ROOT/'akz_stage6/core.py'}:406; exact APPLICATION!A7:G20.
REPRODUCTION: открыть exact приложение в Excel read-only и print preview: 3 строки A4:G6, затем 14 пустых до row 21.
IMPACT: нарушение требуемой компактной рабочей таблицы; числа и полнота выбранных данных корректны.
MINIMAL FIX: скрыть неиспользуемые строки в созданной копии между выбранными данными и итогами.
TARGETED VALIDATION: 3 строки без пустого промежутка в native print preview; 17 сохранены; 18 STOP; итоги и исходные SHA верны.

S6-P1-R2 | MINOR
LOCATION: {ROOT/'akz_stage6/core.py'}:133; {ROOT/'akz_stage6/gui.py'}:83.
REPRODUCTION: исправную pending APPLICATION с SHA pin-шаблона заблокировать CreateFile share=0; вызвать настоящий GUI generate.
Получено: «Не удается открыть файл ... так как формат или расширение этого файла являются недопустимыми».
IMPACT: ошибочная диагностика блокировки как повреждения/неверного формата. Rollback исправен, частичной пары нет.
MINIMAL FIX: понятная ошибка занятости/доступа с файлом и предложением закрыть/повторить, с сохранением rollback.
TARGETED VALIDATION: блокировка второй книги — понятный STOP, no partial/no pending; после снятия блокировки полная пара;
существующий output не изменяется.

17. NOT_RUN

Официальная выдача; реальные назначения АКЗ и реальные измерения/климат/контроль; real climate engine/Td approval;
АКЗ-1/-2/-3; расширение формы сверх 17 строк; Stage 7; production; повторные генерационные E2E Stage 1–5;
полный пользовательский E2E кликами от launcher до публикации; физическая печать; PDF/EXE/deploy.
53 адресных assertions и 16 unit tests не подменяют пользовательский E2E и два найденных acceptance defects.

18. REQUIRED FIX

Только S6-P1-R1 и S6-P1-R2, затем targeted validation из пункта 16.
Рассматриваемые exact XLSX и исходные шаблоны не исправлять внутри этой review; будущая исправленная пара требует новых SHA/evidence.
В этой review fixes не выполнялись.

19. READY FOR USER E2E — NO

Work STOP after review. Stage 7 не начат.
'''
(E/'REVIEW_REPORT.txt').write_text(report,encoding='utf-8')
files={str(p):sha(p) for p in E.rglob('*') if p.is_file() and p.name!='REVIEW_MANIFEST.json'}
manifest={'review':'Stage6 Part1 independent review','created_at':stamp,'verdict':'FAIL',
          'findings':{'BLOCKER':0,'MAJOR':0,'MINOR':2},'runtime_assertions':tests['summary'],'unit_tests':{'PASS':16,'FAIL':0},
          'exact_outputs':before['outputs'],'original_product_and_evidence_unchanged':True,
          'production_reads':0,'production_writes':0,'network_tool_calls':0,
          'original_implementation_manifest':{'path':str(I/'EVIDENCE_MANIFEST.json'),'sha256':sha(I/'EVIDENCE_MANIFEST.json')},
          'guarded_current':guarded,'files':files,'written_last':True,'stop_after_review':True}
save('REVIEW_MANIFEST.json',manifest)
assert all(sha(p)==h for p,h in load(E/'REVIEW_MANIFEST.json')['files'].items())
assert all(sha(p)==h for p,h in guarded.items())
print(json.dumps({'verdict':'FAIL','findings':manifest['findings'],'review_files':len(files),'guarded_files':len(guarded),
    'historical_documents':len(historical_docs),'manifest_sha256':sha(E/'REVIEW_MANIFEST.json'),
    'source_changes':0,'report':str(E/'REVIEW_REPORT.txt')},ensure_ascii=False,indent=2))
