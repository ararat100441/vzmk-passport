import sys
sys.dont_write_bytecode=True
import json,hashlib
from pathlib import Path
OUT=Path(__file__).resolve().parent;ROOT=OUT.parents[1]
def read(p):return json.loads(Path(p).read_text(encoding='utf-8-sig'))
def link(p,label=None):
    p=Path(p).resolve();return '['+(label or p.name)+'](<'+p.as_posix()+'>)'
def local(n,label=None):return link(OUT/n,label)
a=read(OUT/'ACTUAL_HASHES.json');data=read(OUT/'DATA_REVIEW.json');tests=read(OUT/'INDEPENDENT_TARGETED_TESTS.json')
integrity=read(OUT/'SOURCE_INTEGRITY.json');hashes=read(OUT/'HASH_VERIFICATION.json');native=read(OUT/'NATIVE_REVIEW.json')
changes=read(ROOT/'validation/stage7_resume_20261009_175000/CHANGED_FILES.json')
m=read(ROOT/'validation/stage7_resume_20261009_175000/ASSEMBLY_01/STAGE7_SELECTED_TEST_MANIFEST.json')
assert data['passed']==data['total']==108 and tests['passed']==tests['total']==25
assert integrity['status']=='PASS' and not integrity['changes']
assert all(not x['mismatch'] for x in hashes.values())
assert read(OUT/'POINT12_NATIVE_FONT_POSITIONS.json')['equal']
point12=next(c for c in data['checks'] if c['name']=='point12 complete canonical XML')['actual']
lines=['# ВЗМК Этап 7 Независимая проверка продукта', '',
'09.10.2026, Europe/Samara. Проект «Гараж Казахстан», паспорт 9/2026. Только DEV/TEST.', '',
'**FINAL VERDICT: PASS. BLOCKER 0 / MAJOR 0 / MINOR 0. READY FOR USER E2E: YES — только согласованный TEST-комплект.**', '',
'Результат получен независимым чтением кода, повторным вычислением SHA, проверкой фактических данных, новым нативным открытием Word/Excel и собственными отказными тестами. Самопроверка разработчика 46/46 не использована как самостоятельное основание PASS.', '',
'## 1. ACTUAL PRODUCT / OUTPUT HASHES', '',
'SHA `2bfb7d26b54b81a3298dcf23de93ccf76f04945ace281f7dbf50dac426e9f6ea` принадлежит именно выходному **DOCX документа о качестве**, размер 25 429 байт. Это установлено чтением фактического файла и сопоставлением с quality receipt; принадлежность не предполагалась.', '',
'| Артефакт | Фактический SHA-256 |', '| --- | --- |']
for key,label in [('quality','Документ о качестве TEST'),('template','Первый исходник 1/2026'),('manifest','Selected manifest'),
                  ('appendix1','Приложение №1'),('vik','Два акта ВИК в одном DOCX'),('uzk','Новый УЗК SAMPLE_ONLY'),
                  ('akz0','АКЗ №0 TEST'),('akz0-appendix','Приложение к АКЗ №0'),('vom','Принятый M1 vom.py')]:
    lines.append('| '+link(a[key]['path'],label)+' | `'+a[key]['sha256']+'` |')
lines += ['', 'Фактические 12 новых файлов Stage 7:', '', '| Файл | SHA-256 |', '| --- | --- |']
for r in changes['product_files']:lines.append('| '+link(r['path'],str(Path(r['path']).relative_to(ROOT)))+' | `'+a['product:'+Path(r['path']).name]['sha256']+'` |')
lines += ['', 'Сверка: 91 файл принятого checkpoint; 102 файла product; 12 новых файлов; 68 записей implementation evidence — все SHA совпали. '+local('ACTUAL_HASHES.json')+'; '+local('HASH_VERIFICATION.json')+'.', '',
'## 2. INDEPENDENT CODE REVIEW', '',
'Прочитаны полностью Stage 7 core, sources, provenance, word, native, gui, CLI и launcher. Проверены прямые точки вызова принятого Stage 2/5 и относящиеся к текущим четырём маркам правила геометрии. Полный повторный review Этапов 1–6 не выполнялся.', '',
'| Проверенный участок | Результат |', '| --- | --- |',
'| `core.py:20–41`, `sources.py:32–61` | Выбор строки по проекту/номеру, current DEV fields, проверки ВОМ/приложения/АКЗ/УЗК |',
'| `core.py:134–179` | Seal, current membership, обязательные документы, статусы, receipts, run IDs, пары и варианты, текст ссылок |',
'| `core.py:182–245` | Качество последним, новый output, изоляция ВИК, принятые генераторы без правки их кода |',
'| `provenance.py:50–174` | Windows read handles; START до callable; RESULT после фактического output; проверки до/после; успешный receipt после validation |',
'| `word.py:98–196` | Локальные разрешённые замены, сохранение point12/constant cells/numbering/table properties/остальных ZIP parts |',
'| `gui.py:23–56`, `stage7.py` | Явный manifest, TEST-подтверждение, проверка preview и источников, новый путь, read-only Word |', '',
'Установлено 12 добавленных файлов, изменений существующего принятого product — 0. Все 12 отсутствуют в scope предшествующего аудита. Из 427 файлов того более раннего аудита отличается только `vom.py`: старый SHA `236ec7287bc4a9c02cd4e8fff84b60334f5d14b74e00588f0cff90e731857a41`, текущий `c63bbd…954f` уже отдельно принят как M1. Это известное восстановление до Stage 7, а не новый drift. '+local('PREIMPLEMENTATION_DIFF.json')+'.', '',
'Новых дефектов, требующих исправления, не обнаружено. Product не исправлялся.', '',
'## 3. MANIFEST INTEGRITY', '',
'assembly_run_id: `823abd085041423b9776fa83bcc79ce1`. Seal пересчитан; 47 input bindings и 102 code bindings совпали с файлами. Registry SHA: `8e56f68518c70cade9efb09084dd67cd5d72e07e3fdc813b8515030e9cb98e80`; ВОМ SHA: `aae5e5b8c0246c8fe97b3dec9cbcc039bc3e3cedb9d245b474b2f871b90596ae`.', '',
'`read_registry()` / `select_project()` выполнены повторно. 9/2026 найден по проверенному номеру, строка 400 — результат поиска. 10/2026 является следующим номером и в комплект не попал.', '',
'| Запись | Дата | Статус | Связь |', '| --- | --- | --- | --- |']
for r in m['records']:lines.append('| '+r['document_id']+' / '+r['number']+' | '+r['date']+' | '+r['status']+' | '+str(r['source_run_id'] or 'original run неизвестен; не придуман')+' |')
lines += ['', 'Всего шесть логических записей / пять уникальных выбранных файлов. Две записи ВИК используют один файл с разными страницами. АКЗ №1/2/3 отсутствуют. Существующие Приложение №1 и АКЗ №0 явно включены в согласованную assembly; их исторические run IDs не подменены на новые.', '',
'Нативное read-only чтение фактической ВОМ: 104 видимые строки, 397 скрытых исключены, 288 изделий, 15 822,14 кг. Все видимые строки совпали с implementation preview. Приложение №1: 123 строки, те же количество/масса и значения по каждой марке с точностью до округления массы 0,01 кг. '+local('NATIVE_EXCEL_VOM.json')+'; '+local('DATA_REVIEW.json')+'.', '',
'## 4. VIK VALIDATION', '',
'PASS. Новый DOCX содержит два акта: 9/2026 и 9/2026-1, оба от 29.09.2026. Фактические страницы повторно открыты и просмотрены. Номер, договор, объект, КМД, стали и соединения соответствуют текущему DEV. Соединения: «Настил, Площадка, Прогон стеновой, Балка, Прогон, Воротный ригель, Связь, Распорка».', '',
'Побайтная копия DEV в VIK_ISOLATION совпала с текущим registry. Generation record содержит тот же registry SHA, строку 400, exact first-template SHA `1ea022a3079447f21930cb68ec2910a046ac037b80e55fe6ea7ed29644a58fc7` и фактический output SHA. Старый ВИК с несовместимым registry не выбран; его подстановка отклонена отдельным тестом.', '',
'Принятый Stage 2 сохраняет исходную форму ВИК без собственной TEST-плашки. Его TEST-статус задан выбранной изолированной assembly, manifest, receipt и ссылками в качестве. Он не разрешён для отдельной официальной выдачи.', '',
'## 5. UZK / RECORDER VALIDATION', '',
'PASS. Новый УЗК от 01.10.2026; статус SAMPLE_ONLY; `results_confirmed=false`, `official_issue=false`.', '',
'| Марка | Q | Геометрия суммарного контроля, мм |', '| --- | ---: | --- |',
'| Бп2-1 | 8 | 1280 × 5; 1024 × 8,4 |',
'| Пс1-9 | 9 | 3600 × 4 |',
'| Пс1-10 | 15 | 6000 × 4 |',
'| Пс2-16 | 1 | 480 × 3 |', '',
'Итого 4 марки / 33 изделия. Для всех сегментов 5 МГц / 70°. Q независимо сопоставлен с текущим остатком DEV УЗК и видимой ВОМ. Профили прочитаны из текущего справочника; длины пересчитаны из текущей геометрии каталога и Q, затем сверены с физическими ячейками DOCX. ДНО/годен рассматриваются только как образец заполнения.', '',
'| Операция | Generator / operation_id | START → RESULT → VALIDATION → receipt, UTC |', '| --- | --- | --- |']
for r in data['receipts']:
    lines.append('| '+r['operation']+' | `'+r['generator']+'` / `'+r['operation_id']+'` | '+' → '.join(r[x].split('T')[1] for x in ['start','result','validation','receipt'])+' |')
lines += ['', 'Три операции связаны одним assembly_run_id. УЗК/VIK фиксируют 47 inputs / 102 code bindings; качество — 52 inputs / 102 code bindings с выбранными outputs/receipts/manifest. Проверены SHA каждого события, ссылки цепочки, идентичность callable, временной порядок, фактический output, проверки inputs/code до и после. В wrapper нет операции успешного receipt для уже существующего результата: он сначала вызывает зарегистрированный генератор, затем валидирует новый файл.', '',
'Факт исполнения implementation проверен по связанным событиям и пути вызова в неизменном коде; это локальная audit-цепочка, а не криптографическое доказательство против ручной подделки всех файлов на машине. Независимый отрицательный вызов самого `uzk_stage5.core.generate` с испорченным preview завершился PREVIEW_CHANGED, FAILURE и без DOCX/успешного receipt. Отдельно отклонены fabricated leaf без START и попытка выдать receipt для существующего output. Старые receipts не изменены.', '',
'## 6. AKZ №0 VALIDATION', '',
'PASS. Точные actual paths и полные SHA приведены в разделе 1. Выбрана пара из `validation/stage6_fix_r1_r2_20261009_094707/NEW_TEST_9_2026/`: «ТЕСТ АКЗ №0.xlsx» и «ТЕСТ Приложение к АКЗ №0.xlsx». PAIR state COMPLETE_TEST_PAIR, act_id `e9fc12e4433feafd`; SHA обоих файлов совпали с PAIR и ожидаемыми значениями задачи.', '',
'Внутри: номер 9/2026, дата 05.10.2026, один слой Pirocor Protect 057, RAL 7004, требование 80 мкм. Приложение прямо ссылается на этот акт/дату. TEST и «НЕТ ДАННЫХ / ТЕСТ» сохранены; реальные замеры не подтверждены.', '',
'Только Бп2-5, Бп2-18, Бп2-19: три назначенные марки / 3 изделия / 105,26 кг. Принадлежность каждой проверена по видимой ВОМ. Покрытие всех 288 изделий не заявлено; в DOCX есть явное ограничение охвата.', '',
'## 7. QUALITY DOCUMENT DATA', '',
'PASS. Проверены фактические ячейки DOCX относительно current DEV, а не только наличия строк в manifest.', '',
'| Поле | Фактическое значение |', '| --- | --- |',
'| Номер / дата | 9/2026 / 06.10.2026 |',
'| Заказ / договор | 1 / 06/26 от 03.06.2026 |',
'| Заказчик | ООО "СТК ВК" — точное значение DEV |',
'| Объект | Архитектурно-строительные решения здания парковки подвижного состава |',
'| КМД | WAA0007144-AC2-КМД |',
'| Масса | 15 822,14 кг. |',
'| Изготовление | 05.08.2026–06.10.2026 |',
'| Стали | С255, С345 |',
'| Реквизиты | ООО «ВЗМК»; 443076 г. Самара, ул. Крупской, 1 Е |',
'| Подписант | Начальник ОТК; Журавлев А.Н.; пустая подпись; текст М.П. |', '',
'КМ/разработчик и КМД/изготовитель также сверены с J/K/F выбранной строки. Печатное ФИО сохранено; графической подписи и печати нет. На первой странице явно «ТЕСТ / ЧЕРНОВИК. НЕ ДЛЯ ОФИЦИАЛЬНОЙ ВЫДАЧИ». На второй указано отсутствие подтверждения УЗК/АКЗ.', '',
'## 8. POINT 12 IMMUTABILITY', '',
'PASS. Независимо сопоставлены все девять строк 41–49 первой таблицы — от заголовка пункта 12 до фразы хранения сертификатов. Полное canonical XML равенство включает текст, все run/paragraph/cell/row properties, исходное написание, пустые абзацы, объединения. SHA canonical блока у исходника и output: `'+point12+'`.', '',
'Styles, numbering, остальные ZIP parts и общие свойства таблиц/секций не изменены. В native Word блок целиком остаётся на странице 1 и сохраняет исходные переносы/границы. Точное пиксельное равенство PNG не заявляется: дополнительный diff нашёл 177 пикселей внутри одного глифа «л» в слове «толщина», без изменения текста. Повторное открытие показало тот же raster diff; прямое сравнение Word Font, языка, размера, интервалов и X/Y каждого символа фразы подтвердило полное равенство. Изменения формата/геометрии блока не установлены. '+local('POINT12_NATIVE_FONT_POSITIONS.json')+'; '+local('POINT12_NATIVE_RECHECK.json')+'.', '',
'## 9. POINTS 13/14', '',
'PASS. Пункт 13 содержит ровно четыре выбранных акта: ВИК 9/2026 и 9/2026-1 от 29.09.2026; УЗК 9/2026 от 01.10.2026 SAMPLE_ONLY; АКЗ 9/2026 от 05.10.2026 TEST. Пункт 14 содержит эти четыре акта и их выбранное приложение АКЗ, фразу хранения документов и ограничение TEST. Приложение №1 отдельно указано в пункте 2.', '',
'Исторический длинный перечень сертификатов удалён только из разрешённого пункта 14. Каждый текстовый документ имеет реальный выбранный файл. Две ссылки ВИК соответствуют двум актам одного DOCX. Не создано ссылок на АКЗ 1/2/3 или несуществующие документы; число экземпляров не придумано.', '',
'## 10. NATIVE WORD VISUAL REVIEW', '',
'PASS. Самостоятельно открыты текущий quality DOCX, исходник 1/2026, новые ВИК и УЗК. Microsoft Word '+native['quality']['version']+', build '+native['quality']['build']+'; принтер '+native['quality']['printer']+'. `ReadOnly=True`, `OpenAndRepair=False`, макросы и update links отключены, `AddToRecentFiles=False`, закрытие без сохранения. SHA каждого DOCX до/после одинаковый.', '',
'Качество: две физические страницы A4, одна секция, две таблицы; Word list labels 1–14 сохранены; одно исходное MERGEFIELD, без обновления. Нативные EMF получены новым независимым reader и преобразованы GDI в PNG без PDF. Просмотрены обе страницы качества, обе страницы исходника и по две страницы ВИК/УЗК.', '',
'На странице 1 полный пункт 12; на странице 2 пункты 13/14, гарантийная формулировка, должность, ФИО, пустая подпись, дата и М.П. Текст не обрезан, строки не налезают, третьей страницы нет. Переходы строк читаемы; TEST-плашка видима.', '',
local('native_quality/page-1.png','Нативная страница качества 1')+'; '+local('native_quality/page-2.png','Нативная страница качества 2')+'; '+local('NATIVE_REVIEW.json')+'.', '',
'## 11. INDEPENDENT TARGETED TESTS', '',
'**25/25 PASS**. Это собственные тесты этой проверки; suite разработчика не перезапускался. Все изменённые JSON/файлы находятся в `negative_fixtures`, product и реальные источники не изменялись.', '',
'| № | Сценарий | Фактический STOP |', '| --- | --- | --- |']
for i,r in enumerate(tests['tests'],1):lines.append('| '+str(i)+' | '+r['test']+' | '+str(r.get('actual','Windows denied write; original bytes preserved')).split(': ')[0]+' |')
lines += ['', 'Для изменения inputs после preview и существующего output дополнительно проверено `generator_calls=0`; существующие байты сохранены. Ошибки recorder после реального изолированного callable не оставляют успешного receipt. Win32 freeze реально запрещает запись. Отдельно **108/108** независимых сверок содержимого, текущих источников и цепочек SHA — PASS. '+local('INDEPENDENT_TARGETED_TESTS.json')+'; '+local('DATA_REVIEW.json')+'.', '',
'## 12. SOURCE INTEGRITY', '',
'**PASS: 6 431 файла BEFORE / 6 431 AFTER; изменённых, добавленных или удалённых защищённых файлов — 0.** Scope: Pasport и папка паспортов проекта. Включены DEV, код Этапов 1–6, ВОМ, справочники, шаблоны, прежние receipts, validation evidence и существующие паспорта.', '',
'Исключены только собственная evidence-папка этой проверки, dependencies/index/cache/reparse paths, Office lockfiles, PDF/EXE и production registries с точными именами «ВЗМК.xlsx»/«УЗК.xlsm». Production не открывался и не хэшировался. Эти исключения не выдаются за проверенный scope. Отдельно исходные 1 190 bindings implementation BEFORE и AFTER совпали с текущими файлами. '+local('SOURCE_INTEGRITY.json')+'.', '',
'## 13. BLOCKER / MAJOR / MINOR', '',
'BLOCKER 0. MAJOR 0. MINOR 0. Исправлений product не требуется по результатам указанного review. Описанный raster diff одного глифа не сопровождается изменением текста/формата/положения и не квалифицирован как дефект.', '',
'## 14. NOT_RUN', '',
'Ручная приёмка пользователем в живом GUI; новый независимый успешный цикл генерации комплекта; полная регрессия Этапов 1–6; production; запись/пересчёт/синхронизация реестров; списание УЗК; макросы; реальные замеры/подтверждение УЗК и АКЗ; официальная выдача; PDF; EXE; Stage 8. Нативное повторное открытие уже созданных документов и targeted negative tests выполнены и не входят в NOT_RUN.', '',
'Новый codebase index не строился: index_status показал 0 chunks, codebase_context вернул отсутствие результата; использованы точные локальные исходники. Это не ограничило проверку кода.', '',
'## 15. FINAL VERDICT', '',
'**PASS. BLOCKER 0 / MAJOR 0 / MINOR 0.** Вердикт относится к текущим фактическим hashes, выбранной assembly и DEV/TEST-сценарию Stage 7. УЗК/АКЗ и качество всей отгрузки фактически не подтверждены; официальная выдача запрещена.', '',
'## 16. READY FOR USER E2E', '',
'**YES — только DEV/TEST, паспорт 9/2026 и проверенный selected manifest.** Независимая проверка завершена. Исправления, Stage 8 и официальная выдача не запускались. После отчёта STOP.', '']
(OUT/'STAGE7_INDEPENDENT_PRODUCT_REVIEW.md').write_text('\n'.join(lines),encoding='utf8')
visual={'status':'PASS','images_reviewed':['native_quality/page-1.png','native_quality/page-2.png','native_template/page-1.png','native_template/page-2.png','native_vik/page-1.png','native_vik/page-2.png','native_uzk/page-1.png','native_uzk/page-2.png'],
    'point12_xml_equal':True,'native_font_and_positions_equal':True,'pixel_identical':False,
    'pixel_difference':{'count':177,'scope':'one glyph in point12; no font/position/content change'},
    'clipping':False,'overlap':False,'third_page':False,'signed_or_sealed':False}
(OUT/'VISUAL_ASSESSMENT.json').write_text(json.dumps(visual,ensure_ascii=False,indent=2),encoding='utf8')
print(str(OUT/'STAGE7_INDEPENDENT_PRODUCT_REVIEW.md'))
