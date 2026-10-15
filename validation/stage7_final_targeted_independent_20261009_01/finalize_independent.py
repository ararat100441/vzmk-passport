"""Bind the independent result; evidence manifest is written last."""
import sys
sys.dont_write_bytecode=True
from datetime import datetime,timezone
from review_checks import *

h=read(OUT/'HASH_REVIEW.json');t=read(OUT/'INDEPENDENT_TEST_RESULTS.json');n=read(OUT/'NATIVE_VERIFIED_REVIEW.json')
i=read(OUT/'SOURCE_INTEGRITY.json');x=read(OUT/'OOXML_REVIEW.json');chain=read(OUT/'VERIFIED_CHAIN.json')
assert t['status']=='PASS' and t['passed']==24 and i['status']=='PASS' and not i['changes']
assert sha(NEW)==EXPECTED_NEW
def link(p,label=None):return '['+(label or Path(p).name)+'](<'+str(Path(p).resolve()).replace('\\','/')+'>)'
def own(name,label=None):return link(OUT/name,label)
def test(num):return next(v['evidence'] for v in t['tests'] if v['name'].startswith(f'{num:02} '))
save('VISUAL_REVIEW.json',{'status':'PASS','reviewer':'independent review in this chat','method':'Microsoft Word 16.0 native page EMF -> System.Drawing PNG, reviewed both complete pages',
    'word_version':n['new']['word_version'],'word_build':n['new']['word_build'],
    'new_pages':[str(OUT/'native_new'/f'page-{j}.png') for j in (1,2)],
    'manual_pages':[str(OUT/'native_manual'/f'page-{j}.png') for j in (1,2)],
    'first_point12':str(OUT/'native_range_recheck/first/point12.png'),'new_point12':str(OUT/'native_range_recheck/new/point12.png'),
    'new_actual_opened_read_only':True,'reference_snapshots_sha_verified':True,'repair':False,'save':False,'pdf':False,
    'pages':2,'a4':True,'sections':1,'tables':2,'logical_points':14,'clipping':False,'third_page':False,
    'point12_xml_and_pixels_equal_first':True,'point13_14_exact_manual_text_and_effective_format':True,
    'header_warning_prominent':True,'name_role_signature_blank_and_mp_preserved':True,'new_signatures_or_seals':0,
    'accepted_differences':'TEST banner; registry-bound requisites and first-passport point12 retained. Only contract value and point13/14 differ from the prior independently reviewed TEST output.',
    'diagnostic_note':'Initial reviewer range offsets used Content.Text and crossed MERGEFIELD boundaries into point13. Preserved in INITIAL_DIAGNOSTIC files. Corrected only the independent reader to native Find with exact Range.Text assertion. Native point12 now [2036,2702) first and [2077,2743) new, page1, zero differing pixels. No product or source edit.'})
acts=test(1);protect=test(5)
diff_table=['| Файл | Изменение | Проверенный текущий SHA-256 |','|---|---|---|']
for c in h['product_changes']:diff_table.append('| '+link(c['path'],Path(c['path']).name)+' | '+c['kind']+' | `'+c['after']+'` |')
tests_table=['| Независимая проверка | Итог |','|---|---|']
for v in t['tests']:tests_table.append('| '+v['name']+' | '+v['status']+' |')
report=[
'ВЗМК — Этап 7. FINAL TARGETED INDEPENDENT REVIEW / USER E2E FIX',
'',
'**PASS — BLOCKER 0 / MAJOR 0 / MINOR 0. READY FOR FINAL USER E2E: YES, только DEV/TEST.**',
'',
'Проверка выполнена независимо для исправления '+link(FIX/'TARGETED_USER_E2E_REPORT.md')+'. Полный аудит Этапа 7 не повторялся. Самопроверка разработчика 27/27 изучена как входное evidence; её suite не запускалась и не засчитана как независимый тест. Product не изменялся. Все новые файлы этой проверки находятся только в '+str(OUT)+'.',
'',
'1. **ACTUAL NEW DOCX PATH / SHA — PASS**',
'',
link(NEW,'Новый выходной Документ о качестве (ТЕСТ)')+'.',
'',
'SHA-256: `'+sha(NEW)+'` — совпадает с ожидаемым SHA пользователя, с новым RESULT и новым receipt. Сверено до и после native Word. Прежний reviewed DOCX: '+link(OLD/'QUALITY_TEST/9_2026 - Документ о качестве (ТЕСТ).docx')+', SHA `'+read(OUT/'ACTUAL_HASHES.json')['old']['sha256']+'`. Пути и SHA различны; прежний файл совпадает со снимком предыдущей независимой проверки.',
'',
'Оба оригинала сверены с `stage7_quality_audit_20261009/source_before.json`: первый '+link(FIRST)+' — `'+sha(FIRST)+'`; ручной '+link(MANUAL)+' — `'+sha(MANUAL)+'`. Текущие SHA равны audit и не изменились за review. '+own('ACTUAL_HASHES.json')+'; '+own('HASH_REVIEW.json')+'.',
'',
'2. **EXACT PRODUCT DIFF — PASS**',
'',
'Независимо восстановлен '+own('INDEPENDENT_EXACT_PRODUCT_DIFF.diff','полный diff: три изменения и один добавленный файл')+'. Содержимое diff трёх изменённых файлов совпадает с supplied EXACT_PRODUCT_DIFF; новый helper включён целиком. BEFORE-копии совпадают с SHA предыдущего независимого review, а не только с заявлением автора исправления.',
'',
*diff_table,
'',
'word.py: чистые названия актов из type/number/date, последовательный перечень через запятую, точная фраза хранения, без AKZ-app/SCOPE в Word, снятие одного лишнего начального № в значении договора. core.py: frozen manifest проверяется с точными историческими переходами, новый receipt связывает текущий код и архивные inputs. provenance.py: те же строгие переходы при проверке прежних producer receipts. presentation_revision.py: ровно три известные old→new пары SHA и обязательные архивные BEFORE-файлы; неизвестные SHA и прочие файлы проходят строгую проверку. Отрицательные fixture-проверки обхода пройдены. Других product changes и дополнительных source-файлов в bound product directories нет.',
'',
'3. **POINT 12 VERIFICATION — PASS**',
'',
'Диапазон определён самостоятельно по заголовкам пунктов 12/13. Все строки 41–49 полного canonical OOXML совпадают с первым паспортом 1/2026: текст, пробелы, абзацы, pPr/rPr, шрифты, стили, списки, numPr, границы, tcPr, объединения и свойства строк. Все прочие ZIP parts, включая styles/numbering/theme, побайтно сохранены. Сохранены Pirocor Protect 057, «толщина слоя не менее 80 мкм.», «Пескоструйная обработка. Степень очистки 2» и историческое «(степень очистки от окалины ржавчины по ГОСТ 9.402)».',
'',
'Native Find: первый Word Range [2036,2702), новый [2077,2743), оба страница 1, одинаковые текст и координаты 620.6500–733.6500 pt. Независимый прямой Range EMF→PNG: '+str(protect['native_range_png_dimensions'])+', **0 отличающихся пикселей**. Изолированное изменение pPr отвергнуто guard. '+own('POINT12_FIRST_CANONICAL.xml')+'; '+own('POINT12_NEW_CANONICAL.xml')+'; '+own('native_range_recheck/first/point12.png','первый native point12')+'; '+own('native_range_recheck/new/point12.png','новый native point12')+'.',
'',
'4. **POINT 13 VERIFICATION — PASS**',
'',
'Акт ВИК № 9/2026 от 29.09.2026 г., Акт ВИК № 9/2026-1 от 29.09.2026 г., Акт УЗК № 9/2026 от 01.10.2026 г., Сертификат качества антикоррозийной обработки № 9/2026 от 05.10.2026 г.',
'',
'Один абзац. Текст, порядок, запятые, пробелы и свойства каждого видимого символа/абзаца совпадают с ручным 9/2026. Разбиение на OOXML runs отличается, но формат всех символов одинаков; native Word подтверждает те же шрифт, размер, paragraph format, стиль, страницу и положение. Служебных меток нет. Четыре записи соответствуют verified manifest. Динамика независимо проверена с четырьмя различными номерами и датами 2032 года и намеренно неверными display_title; presenter использует record.number/date. Это in-memory presenter test, не новая разрешённая assembly будущего паспорта.',
'',
'5. **POINT 14 VERIFICATION — PASS**',
'',
'Первый абзац — тот же точный перечень четырёх актов, что в пункте 13. Второй абзац: «Документы о качестве на металлопрокат, сварочные материалы и лакокрасочные материалы хранятся на предприятии». Текст и оформление совпадают с ручным 9/2026. Пояснение «(перечень документов с указанием числа экземпляров)» сохранено целиком; количество экземпляров не добавлено.',
'',
'Нет отдельной записи AKZ-app, служебных статусов, описания TEST-комплекта, заявления об отсутствии подписи, количества тестовых марок или исторического перечня сертификатов. AKZ и AKZ_APPENDIX остаются обязательной парой в manifest/receipt inputs: проверены оба файла/SHA, pair_id `'+test(4)['pair_id']+'`, parent_id и variant=0. Отсутствующий appendix record/file и неправильный pair_id приводят к STOP до writer.',
'',
'6. **CONTRACT NUMBER — PASS**',
'',
'В Word: «Договор № 06/26 от 03.06.2026». Постоянная метка «Договор №» и значение совпадают с ручным образцом, двойного № нет. В сравнении с прежним reviewed TEST DOCX изменены ровно три ячейки: значение договора, пункт 13 и пункт 14. Все остальные ячейки, реквизиты, таблицы/секции/нумерация/MERGEFIELD и непрозрачные package parts сохранены. '+own('OOXML_REVIEW.json')+'.',
'',
'7. **TEST STATUS / OFFICIAL BLOCK — PASS**',
'',
'В заголовке заметно: «ТЕСТ / ЧЕРНОВИК. НЕ ДЛЯ ОФИЦИАЛЬНОЙ ВЫДАЧИ.». Manifest mode=DEV_TEST_ONLY, official_issue=false. Новый receipt status=TEST, state=COMPLETE_VALIDATED_TEST, official_issue=false, results_confirmed=false. УЗК SAMPLE_ONLY, official_confirmed=false. АКЗ №0: Бп2-5, Бп2-18, Бп2-19, 3 изделия из отгрузки 288, actual_measurements_confirmed=false. Удаление административного текста из пунктов 13/14 не меняет metadata/gates.',
'',
'Отрицательные official manifest, receipt, writer/presenter и повышение статуса УЗК отвергнуты; writer для невалидных manifest не вызывался. Пустая строка подписи, должность, ФИО и М.П. сохранены из прежнего TEST DOCX; изображений подписи/печати и новых shapes/inline shapes нет. Официальная выдача остаётся заблокирована.',
'',
'8. **NATIVE WORD VISUAL REVIEW — PASS**',
'',
'Microsoft Word '+n['new']['word_version']+', build '+n['new']['word_build']+'. Новый исходный output открыт непосредственно с ReadOnly=True, OpenAndRepair=False, без сохранения; Close(SaveChanges=0). Изолированный owned Word закрыт, пользовательская Word-сессия не затронута. Оригиналы references прочитаны через SHA-проверенные побайтные копии в независимом evidence; оригиналы не сохранялись.',
'',
'Две физические A4 страницы 595.30×841.90 pt, одна секция, две таблицы, логические пункты 1–14. Обе страницы нового файла самостоятельно просмотрены рядом с обеими страницами ручного 9/2026. Пункты 13/14 расположены и оформлены как в ручном образце; обрезки, перекрытий, потери текста и третьей страницы нет. Разрешённые TEST/registry differences сохранены, пункт 12 соответствует именно первому паспорту. Подпись/ФИО/М.П. и пустое место подписи сохранены. '+own('VISUAL_REVIEW.json')+'; '+own('NATIVE_VERIFIED_REVIEW.json')+'; '+own('NEW_MANUAL_BOTH_PAGES.png','сравнение обеих страниц')+'.',
'',
'В начальном независимом диагностическом запуске plain-text offsets пересекли MERGEFIELD и захватили часть пункта 13 («1/» против «9/»). Исправлен только reader независимой проверки: native Find с проверкой точного Range.Text. Исходная диагностика сохранена в INDEPENDENT_TEST_RESULTS_INITIAL_DIAGNOSTIC.json и native_*/WORD.json; окончательное evidence — NATIVE_VERIFIED_REVIEW.json и native_range_recheck. Это ошибка диапазона reviewer fixture, дефект product не установлен.',
'',
'9. **INDEPENDENT TARGETED TESTS — PASS: 24/24**',
'',
'Все десять запрошенных тем покрыты отдельной suite '+own('independent_tests.py')+'. Дополнительно выполнены 17 отрицательных cases. Все изменённые manifest/receipt/code fixtures созданы только в независимой папке. Новые ВИК/УЗК/АКЗ и quality не генерировались. '+own('INDEPENDENT_TEST_RESULTS.json')+'; '+own('NEGATIVE_FIXTURES.json')+'.',
'',
*tests_table,
'',
'10. **MANIFEST / RECEIPT — PASS**',
'',
link(MANIFEST,'Исходный frozen manifest')+' — SHA `'+sha(MANIFEST)+'`, assembly_run_id `'+chain['assembly_run_id']+'`. Все шесть записей, files/SHA, membership, статусы, producer receipts и АКЗ-пара проверены. Manifest не переписан/reseal на source не выполнялся.',
'',
link(RECEIPT,'Новый quality receipt')+' — SHA `'+sha(RECEIPT)+'`, operation_id `'+chain['operation_id']+'`, generator=stage7_quality.word.write_quality. START→RESULT→VALIDATION проверены по SHA, идентификаторам, inputs/code, output path/SHA, PASS и времени. Receipt связывает '+str(chain['code_files'])+' текущих code files, исходный manifest, все выбранные outputs и три архивных BEFORE-файла. Старые producer chains сохранены и проходят проверку через строго заданные переходы. '+own('VERIFIED_CHAIN.json')+'.',
'',
'11. **SOURCE INTEGRITY — PASS**',
'',
'Самостоятельно перехешированы **1264/1264 protected files**, BEFORE=AFTER=CURRENT, несовпадений 0. Текущий bound product **103/103** совпадает с проверенными SHA до/после review. Изменения Stage 1–6: 0. В protected scope сохранены 66 DOCX, 3 прежних Stage7 operation receipts, 39 файлов прежней assembly, source registries/VOM/catalog/profiles и исходный frozen manifest. Новый output имеет отдельный путь и не перезаписал старый. Объединённый независимый scope из protected/product/fix-evidence содержит 1354 файла: изменений за review 0. '+own('SOURCE_BEFORE.json')+'; '+own('SOURCE_AFTER.json')+'; '+own('SOURCE_INTEGRITY.json')+'.',
'',
'Product/source writes: 0. Генераций документов: ВИК=0, УЗК=0, АКЗ=0, QUALITY=0. Отдельно выполнен один отрицательный вызов official guard в quality writer, отвергнутый до чтения шаблона/создания output. Production registry не открывался; УЗК не списывался. Новые артефакты созданы только в независимом evidence directory.',
'',
'12. **FINDINGS**',
'',
'BLOCKER 0 / MAJOR 0 / MINOR 0. Подтверждённых дефектов в targeted scope нет.',
'',
'13. **NOT_RUN**',
'',
'Полный повторный review Stage 7; full regression Stage 1–6; повторные генераторы ВИК/УЗК/АКЗ/quality; developer 27-test suite; полная GUI E2E; финальная ручная приёмка пользователем; фактические контрольные измерения/официальная выдача; production registry access/writes и списание УЗК; PDF/EXE/build/deploy; Stage 8. Индексные index_status/codebase_context возвращали INTERNAL_ERROR: использованы точечные чтения известных файлов и diff; индекс не перестраивался.',
'',
'14. **FINAL VERDICT: PASS**',
'',
'Точечное исправление и связанные с ним риски независимо проверены. Ограничения DEV/TEST сохранены; отсутствие подтверждённых УЗК/АКЗ результатов не превращено в официальный выпуск.',
'',
'15. **READY FOR FINAL USER E2E: YES**',
'',
'Готов к финальной пользовательской приёмке именно нового output с SHA `'+EXPECTED_NEW+'`. Это не разрешение на официальную выдачу. После отчёта STOP. Stage 8 не запускался.',
]
(OUT/'FINAL_TARGETED_INDEPENDENT_REVIEW.md').write_text('\n'.join(report)+'\n',encoding='utf8')
# Final source rehash after all reads/Office/test activity, before manifest-last.
after()
assert read(OUT/'SOURCE_INTEGRITY.json')['status']=='PASS'
evidence={str(p):{'sha256':sha(p),'bytes':p.stat().st_size} for p in sorted(OUT.rglob('*')) if p.is_file() and p.name!='INDEPENDENT_EVIDENCE_MANIFEST.json'}
manifest={'schema':'STAGE7_FINAL_TARGETED_INDEPENDENT_EVIDENCE_V1','created_at':datetime.now(timezone.utc).isoformat(),
    'review':'FINAL TARGETED INDEPENDENT REVIEW USER E2E FIX','verdict':'PASS','findings':{'BLOCKER':0,'MAJOR':0,'MINOR':0},
    'ready_for_final_user_e2e':True,'dev_test_only':True,'official_issue':False,
    'new_docx':{'path':str(NEW),'sha256':sha(NEW)},'source_scope_count':1354,'protected_count':1264,'product_count':103,
    'tests':{'total':24,'passed':24,'failed':0,'negative_cases':17},'evidence':evidence}
with (OUT/'INDEPENDENT_EVIDENCE_MANIFEST.json').open('x',encoding='utf8') as stream:json.dump(manifest,stream,ensure_ascii=False,indent=2)
for p,v in evidence.items():assert sha(p)==v['sha256']
print(json.dumps({'report':str(OUT/'FINAL_TARGETED_INDEPENDENT_REVIEW.md'),'manifest':str(OUT/'INDEPENDENT_EVIDENCE_MANIFEST.json'),'evidence_files':len(evidence),'verdict':'PASS'},ensure_ascii=False))
