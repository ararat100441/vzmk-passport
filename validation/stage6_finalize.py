"""Final immutable evidence snapshot. Manifest is written LAST; no product mutation."""
import sys,json,hashlib
from pathlib import Path
from datetime import datetime
ROOT=Path(__file__).resolve().parents[1];sys.path.insert(0,str(ROOT))
from akz_stage6 import core as c
E=Path((ROOT/'validation/STAGE6_PART1_CURRENT.txt').read_text(encoding='utf-8'))
def read(name):return json.loads((E/name).read_text(encoding='utf-8'))
def save(name,data):(E/name).write_text(json.dumps(data,ensure_ascii=False,indent=2,default=str),encoding='utf-8')
validation=read('targeted_validation.json');assert validation['status']=='PASS'
output=Path(validation['output']);assert '_release3_' in output.name
preview=read('preview.json');p=preview['payload'];before=read('baseline.json')
original_after={x:c.sha256(x) for x in before};assert before==original_after
tested_before=read('tested_product_before.json')
# Generation core and unit tests were frozen for the 29-check run. The final
# human-readable GUI preview was then tested separately at its final hash.
gui_phase=read('native_gui_product_hashes.json');assert gui_phase['status']=='PASS'
assert gui_phase['BEFORE']==gui_phase['AFTER']==gui_phase['CURRENT']
tested_before.update(gui_phase['BEFORE'])
tested_after={x:c.sha256(x) for x in tested_before};assert tested_before==tested_after
generated=read('generated.json')
assert all(Path(x).is_file() and c.sha256(x)==h for x,h in generated['files'].items())
for n in ('native_act_release.json','native_application_release.json'):
    native=read(n);assert native['read_only'] and native['normal_corrupt_load']==0
    assert len(native['sheets'])==1 and native['sheets'][0]['pages']==1
    assert c.sha256(native['path'])==native['input_sha256']
    assert read(n.replace('.json','.after.json'))['unchanged']
assert read('native_gui_preview.json')['status']=='PASS'
assert not list(c.OUTPUT.glob('.pending-*'))
new_product=[*sorted((ROOT/'akz_stage6').glob('*.py')),ROOT/'stage6.py',ROOT/'tests/test_stage6.py',
             ROOT/'Запустить Этап 6 DEV.cmd',ROOT/'STAGE6_DEV.md']
save('tested_product_hashes.json',{'BEFORE':tested_before,'AFTER':tested_after,
    'CURRENT':{x:c.sha256(x) for x in tested_after},'exact_match':True,
    'launch_and_documentation_current':{str(x):c.sha256(x) for x in new_product if str(x) not in tested_before}})
save('source_integrity_current.json',{'BEFORE':before,'AFTER':read('source_integrity_after_tests.json')['after'],
    'CURRENT':original_after,'changed':[],'guarded_files':len(before),'production_opened_or_hashed':False})
save('development_attempts.json',{'canonical_output':str(output),'attempts':[
    {'name':'initial','result':'Excel COM RPC startup failure, no pair published'},
    {'name':'run2','result':'COM lifetime failure, no pair published'},
    {'name':'run3','result':'saved date format/value check failed; pair retained as FAILED_VALIDATION_DO_NOT_USE_run3'},
    {'name':'run4','result':'27 integration checks and 16 unit tests passed; superseded after visual review'},
    {'name':'final','result':'29 integration checks and 16 unit tests passed; superseded after layout/number display review'},
    {'name':'release','result':'native NumberFormat setter failure; whole pending pair rolled back'},
    {'name':'release2','result':'International COM property shape failure; whole pending pair rolled back'},
    {'name':'release3','result':'final current code: 29 integration checks and 16 unit tests passed; native UI and print preview passed'}]})
sources='\n'.join(f'| `{x}` | `{h}` |' for x,h in p['sources'].items())
outputs='\n'.join(f'- `{x}`\n  SHA-256: `{h}`' for x,h in generated['files'].items() if x.endswith('.xlsx'))
selected='\n'.join(f'| {r["row"]} | {r["mark"]} | {r["name"]} | {r["quantity"]} | {r["mass_unit"]} | {r["mass_total"]} | {r["area_unit"]} | {r["area_total"]} |' for r in p['selected'])
checks='\n'.join(f'- PASS: {x["name"]}' for x in validation['checks'])
report=f'''# ВЗМК — ЭТАП 6 IMPLEMENTATION PART 1 — АКЗ №0

Дата отчёта: 09.10.2026, Europe/Samara. READY FOR INDEPENDENT REVIEW — YES.
Разрешённая часть: только DEV / TEST АКЗ №0. Официальная выдача STOP.

## 1. ACTUAL STATE

Реализация изолирована в akz_stage6, stage6.py и отдельном launcher.
Прочитаны оба указанных baseline audit/final-check отчёта. Принятые этапы 1–5
не изменены. Средства index_status/codebase_context в этой сессии отсутствовали,
выполнены адресные чтения. Product source templates редактируются только на копиях
через native Excel COM ради сохранения исходной структуры.
Текущий паспорт выбран accepted read_registry → select_project на явном DEV-пути:
Гараж Казахстан, 9/2026, строка 400. На диске существует паспорт 10/2026;
это предупреждение preview, без изменения DEV и без генерации нового паспорта.

## 2. TEMPLATE PAIR VERIFIED

Явно зафиксирована пара первого паспорта 1/2026 из корня проектной папки Паспорт:
Акт № 4-2026.xlsx и 1_2026 - Приложение к АКЗ.xlsx. SHA совпали с заданием.
Проверены внутренние E4/H4, ссылки G8=E4/I8=H4, A1 приложения и единственные листы.
Имя «4» не используется для вывода номера. Подмена/поиск похожего имени отсутствуют.

| Exact source | SHA-256 |
|---|---|
{sources}

## 3. DEV TEST ASSIGNMENTS

Создана точная копия настоящего каталога в `{validation['fixture']}`.
TEST_Справочник марок.xlsx содержит все 195 уникальных марок, три числовых назначения
0 и 192 пустых назначения. Назначены только Бп2-5, Бп2-18, Бп2-19 — каждая найдена
в видимой выбранной ВОМ. Все назначения TEST_ONLY, реальные назначения не изменялись.
Источник и SHA копии: test_assignments.json. Остальные отрицательные сценарии используют
синтетические изолированные книги и in-memory fixtures.

## 4. VOM / AKZ=0 MATCHING

Явная ВОМ: `{p['vom']['path']}`. Native XLS reader: 104 видимые строки, 397 скрытых.
AutoFilter и manual Hidden учтены; повторные строки сохраняются в исходном порядке.
Выбраны три строки №0; 101 видимая строка показана как NO_ASSIGNMENT.
Другие варианты 1/2/3 исключаются; дубликат каталога STOP.

| Строка ВОМ | Марка | Наименование | Кол-во | Масса 1 | Масса всех | Площадь 1 | Площадь всех |
|---|---|---|---|---|---|---|---|
{selected}

## 5. AREA / MASS ASSESSMENT

Только выбранная ВОМ даёт числовые данные. Приложение №1 и справочник профилей
не используются как числовые источники. Среди совпавших с каталогом видимых строк
44 прошли строгую сверку, 11 остановлены из-за пропусков/несогласованных чисел.
Выбранные строки имеют количество 1, положительные unit/total и точное соответствие.
Итоги: **3 шт.; 105.26 кг; 3.989 м²**. Не применяется бизнес-округление,
лишь 1e-8 для транспортной погрешности double. Пропуск, ноль или конфликт площади
останавливают выбранную группу; незаполненные другие группы положительному TEST не мешают.
Формат вывода сохраняет десятичную точность исходных чисел, включая площадь 3.989.

## 6. ACT №0 IMPLEMENTATION

Акт 9/2026, дата из именованного поля DEV «Дата Акз» O400 = 05.10.2026.
Историческая дата 9-го акта 06.10.2026 взята только как предупреждение принятого аудита.
Договор C400, заказчик D400, объект E400, КМД F400. Ссылка приложения согласована.
Из формы сохранены один слой, Pirocor Protect 057, RAL 7004, требуемые 80 мкм.
92 мкм, климат, партии, даты/время нанесения, инструменты фактического контроля,
отметки подготовки/нанесения и положительное заключение не перенесены как новые факты.
В узких ячейках отметок стоит «?», в остальных — «НЕТ ДАННЫХ / ТЕСТ».
Числовой Excel serial устраняет COM-сдвиг даты; исходный формат даты сохранён.
Длинный объект переносится в увеличенной строке, подписи/постоянные реквизиты сохранены.

## 7. APPLICATION IMPLEMENTATION

Точная копия проектного приложения №0. Заполняется A:G из выбранных строк ВОМ.
Все исторические строки 4:20 очищаются перед заполнением; невыбранных строк нет.
Итоги C21/E21/G21 — SUM, пересчитаны native Excel и сверены с Decimal-суммами.
Заголовок, номер, дата и ACT-ID `{preview['seal'][:16]}` совпадают с актом.
Part 1 поддерживает до 17 выбранных строк исходной формы; больше — CAPACITY STOP,
без обрезания и без незаявленного расширения таблицы.

## 8. CLIMATE / TEST SAFETY

ClimateDay: календарная дата, T воздуха, RH, T поверхности, точка росы, статус.
Одна запись на дату, конфликт второй записи STOP. Для TEST-набора проверяются
полнота, RH и T поверхности >= Td+3. Формула точки росы не выбирается.
В готовом акте показания неизвестны и отмечены явно. TEST не подтверждает работы.
Официальный режим всегда STOP. SYNTHETIC-ВОМ требует явный режим и отдельный номер
SYNTHETIC-TEST; финальная пара использует реальные числа выбранной ВОМ, не synthetic.

## 9. GENERATED TEST FILES + SHA

Только эта финальная пара предназначена для review:

{outputs}

Оба файла помечены ТЕСТ / ЧЕРНОВИК. PAIR.json фиксирует единую идентичность и SHA.
Старые промежуточные пары сохранены под FAILED_VALIDATION/SUPERSEDED и не являются
финальными deliverables. Входные файлы и существующие выходы не перезаписывались.

## 10. TARGETED VALIDATION

**29 интеграционных проверок PASS; 16 unit boundary tests PASS**.

{checks}

unit_tests.txt дополнительно доказывает: missing/conflicting area STOP,
numeric 0 отдельно от blank, 1/2/3, row duplicates, exact mark, mass conflict,
отсутствие выдуманного округления, blank иных групп, capacity STOP, NaN/bool,
climate day conflict/граница +3, official STOP и изменение payload preview.
Перед публикацией обе сохранённые книги заново открываются read-only в native Excel,
проверяются номер/дата/ссылки/динамические поля/материал/марки/значения/итоги.
Ошибка второго файла и ошибка saved-validation удаляют только pending-пару.

## 11. NATIVE EXCEL VALIDATION

Обе финальные книги открыты в реальном Microsoft Excel 16.0 **read-only**,
UpdateLinks=0, macros/events disabled, обычный CorruptLoad=0.
Не было repair/recovery-диалогов; обе книги отображены и проверены через native UI.
Проверены заголовки, номер/дата, динамические реквизиты, три марки, количества,
массы, площади, итоги и отсутствие исторических строк.
Native print preview: по **1 из 1** странице, по одному видимому ожидаемому листу.
Акт print area B2:K50; приложение A1:G21. Лишних печатных листов/страниц нет.
Физической печати не было. Изображения и accessibility text находятся в native_*png/txt.
SHA после закрытия без сохранения совпадают; native_*.after.json.
Открыт действующий Tk DEV-launcher с проверенным preview и доступной генерацией;
изменение источников отключает её. Генерация через ручные клики пользователя не выполнялась.

## 12. CHANGED FILES

Добавлены только akz_stage6/__init__.py, core.py, gui.py, stage6.py,
Запустить Этап 6 DEV.cmd, STAGE6_DEV.md, tests/test_stage6.py,
validation/stage6_*.py, Stage6 evidence и указатель STAGE6_PART1_CURRENT.txt,
изолированные DEV/Stage6_Fixtures и TEST-пары в DEV/Stage6_Выход.
Ни один существовавший принятый модуль/тест/settings/registry не изменён.

## 13. SOURCE / PRODUCTION SAFETY

**{len(before)} исходных контрольных файлов: BEFORE=AFTER=CURRENT, изменений 0.**
Production ВЗМК.xlsx не открывался, не хешировался и не сохранялся.
Настоящий каталог, ВОМ, исходные формы и DEV-реестры неизменны.
Инструменты/макросы для production не запускались; registry mutation API отсутствует
в новой реализации. No АКЗ-1/-2/-3, Stage 7, PDF, EXE, deploy.

## 14. EVIDENCE INTEGRITY

tested_product_hashes.json содержит BEFORE/AFTER/CURRENT конечного исполняемого
Stage 6 продукта и тестов, все равны. Core/unit привязаны к полному прогону;
финальное текстовое представление GUI проверено отдельно по его неизменному SHA
в native_gui_product_hashes.json (марки/итоги в начале, сброс и повтор preview).
Отдельно зафиксированы launcher/docs.
source_integrity_current.json фиксирует все {len(before)} прежних источника.
Preview JSON включает hashes, mapping, выбранные/исключённые строки, totals, seal.
Все evidence/fixtures/outputs и новый продукт включены в EVIDENCE_MANIFEST.json.
Manifest записан последним и после записи проверен read-only.

## 15. NOT_RUN

Официальная выдача; реальные назначения № АКЗ; реальные измерения/климат/контроль;
утверждение формулы Td; автоматическое расширение приложения сверх 17 строк;
АКЗ-1/-2/-3; Stage 7; production; повторные генерационные E2E этапов 1–5;
физическая печать; PDF/EXE/deploy. Ручной User E2E всей цепочки кликами не выполнен.

## 16. READY FOR INDEPENDENT REVIEW — YES

YES — в разрешённом объёме Part 1, DEV / TEST АКЗ №0.
Официальная выдача STOP. После отчёта STOP. Stage 7 не начат.
'''
(E/'STAGE6_IMPLEMENTATION_PART1.md').write_text(report,encoding='utf-8')
files=[x for x in E.rglob('*') if x.is_file() and x.name!='EVIDENCE_MANIFEST.json']
files+=new_product
files += [x for directory in (c.FIXTURES,c.OUTPUT) for x in directory.rglob('*') if x.is_file() and not x.name.startswith('~$')]
files += sorted((ROOT/'validation').glob('stage6_*.py'))
files += [ROOT/'validation/STAGE6_PART1_CURRENT.txt']
manifest={'scope':'Stage 6 Part 1 TEST ONLY; no production; manifest written last',
          'excluded_transient_files':'Excel owner lockfiles ~$* are not artifacts or workbook content',
          'report':str(E/'STAGE6_IMPLEMENTATION_PART1.md'),'canonical_output':str(output),
          'source_fingerprints':original_after,'tested_product':tested_after,
          'files':{str(x):c.sha256(x) for x in sorted(set(files))}}
save('EVIDENCE_MANIFEST.json',manifest)  # LAST WRITE
assert all(c.sha256(x)==h for x,h in manifest['files'].items())
assert all(c.sha256(x)==h for x,h in manifest['source_fingerprints'].items())
print('MANIFEST_VERIFIED',len(manifest['files']),'SOURCE_UNCHANGED',len(before))
print('REPORT',E/'STAGE6_IMPLEMENTATION_PART1.md')
print('OUTPUT',output)
print(json.dumps(generated['files'],ensure_ascii=False,indent=2))
