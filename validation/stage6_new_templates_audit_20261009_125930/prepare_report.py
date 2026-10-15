"""Build audit report and non-executable implementation contract from captured evidence."""
from pathlib import Path
import json, hashlib
OUT=Path(__file__).resolve().parent
ROOT=OUT.parents[1]
def read(name):return json.loads((OUT/name).read_text(encoding='utf-8'))
def write(name,value):(OUT/name).write_text(json.dumps(value,ensure_ascii=False,indent=2,default=str),encoding='utf-8')
static=read('static_workbooks.json'); native=read('native_excel.json'); inventory=read('selector_inventory.json'); before=read('sources_before.json')
folder=Path(inventory['folder'])
common={'number':'E4','date':'H4','contract':'E5','customer':'E6','object':'E7','cipher':'B44','appendix_number_formula':{'G8':'=E4'},'appendix_date_formula':{'I8':'=H4'},'commission_names':['E46','E47'],'signature_date':'D49','signature_boxes':['F49:G49','I49:K49'],'signature_captions':['F50','I50']}
variants={}
for n,cols,layer_numbers,total in [(1,['E'],['E17'],66),(2,['E','I'],['E17','H17'],52),(3,['E','H','J'],['E17','H17','J17'],99)]:
    act=folder/f'АКЗ-{n}.xlsx'; app=folder/f'Приложение к АКЗ-{n}.xlsx'
    s=next(b for b in static if b['path']==str(act))['sheets'][0]
    layers=[]
    for index,(col,num) in enumerate(zip(cols,layer_numbers),1):
        layers.append({'ordinal':index,'ordinal_cell':num,'column':col,'material':{'cell':col+'18','value':s['cells'][col+'18']['value']},'color':{'cell':col+'21','value':s['cells'][col+'21']['value']},'required_thickness':{'cell':col+'25','raw':s['cells'][col+'25']['value'],'status':'UNRESOLVED_CUMULATIVE_REQUIREMENT' if n==2 and index==2 else 'EXPLICIT_NUMERIC_MICRONS'},'historical_fields':{key:col+str(row) for key,row in [('paint_batch',19),('hardener_batch',20),('application_date',22),('start_time',23),('end_time',24),('row26_ambiguous_total',26),('wet_thickness',27),('dry_thickness',28),('climate_day',30),('surface_c',31),('air_c',32),('humidity',33),('dew_point_c',34)]}})
    variants[str(n)]={'status':'PARTIAL','act_path':str(act),'act_sha256':before[str(act)],'act_sheet':'4','act_fields':common,'layers':layers,'application':{'path':str(app),'sha256':before[str(app)],'sheet':'Приложение №1','header':'A1','object':'A2','column_headers':'A3:G4','body_first':5,'body_last':total-1,'capacity':total-5,'total_row':total,'columns':dict(zip('ABCDEFG',['mark','name','quantity','mass_unit','mass_total','area_unit','area_total'])),'formulas':{col+str(total):f'=SUM({col}5:{col}{total-1})' for col in ('C','E','G')},'print_area':f'A1:G{total}','print_title_rows':'1:4'},'native':[{k:v for k,v in x.items() if k!='sheets'}|{'sheets':[{k:v for k,v in y.items() if k not in ('values','formulas')} for y in x['sheets']]} for x in native if x['variant']==n]}
contract={'kind':'AUDIT_ONLY_NOT_IMPLEMENTATION','state':'PARTIAL','variant_numbering':{'0':'9/2026','1':'9/2026-1','2':'9/2026-2','3':'9/2026-3'},'rules':{'variant0':'Existing implementation, pins, output artifacts and R1/R2 behavior remain unchanged','coating_requirements':'Selected pinned template only; never suffix or catalog paint column','quantity_mass_area':'Explicit selected current VOM only; preserve original row order and repeated rows; missing area stays missing','act_date':'DEV named header Дата Акз, current row, currently O400 = 2026-10-05','climate':'One confirmed record per calendar date shared across all variants and layers with that application date; no random formulas','dew_point':'ALGORITHM_NOT_SELECTED; do not endorse source formula','historical_results':'Never reuse template measurements, times, batches, climate, instrument attestations or signatures as new evidence','pair_identity':'Exact paths are candidate pairs supplied by user; source workbook contents conflict; generated-copy headers must be explicitly rebound to current act+suffix and DEV date','capacity':'Fixed verified template body bounds; N > capacity stops before copy/publication; no automatic row insertion','publication':'Do not weaken accepted UUID ownership, exact file allowlist, durable pending ledger, primary error preservation, two-file validation, atomic rename; variants require their own exact allowlist and pair identity; leave variant0 path unchanged'},'variants':variants,'unresolved':['Variant2 second layer has cumulative text ∑1+2=180, not an explicit per-layer requirement; do not silently derive 80','Variant1 color is серый without numeric RAL; variants2/3 final color is numeric 7015 without explicit RAL label','Application headers identify quality documents 4/2026, 5/2026, 6/2026, not matching source acts; authorize exact generated-copy rebinding before implementation','Variant2 visible service Лист1 must have an explicit output/print policy; do not publish historical helper content accidentally','Row26 total vs measured and row28 incremental vs cumulative semantics need confirmation before official calculations','No confirmed dew-point/climate-control algorithm or fresh official measurements; keep official issuance STOP']}
write('IMPLEMENTATION_CONTRACT.json',contract)
report=[]
def add(t):report.append(t.strip()+'\n')
add('''# ВЗМК — Stage 6 — read-only audit новых АКЗ-1/-2/-3

## 1. ACTUAL STATE

Аудит выполнен 09.10.2026. Implementation НЕ начат. Вызваны принятые `read_registry()` и `select_project()` до поиска проектных файлов: Гараж Казахстан, первый паспорт 1/2026, строка 353; последний в DEV 9/2026, строка 400. Папка первого паспорта — непосредственно `Гараж казахстан\\Паспорт`, без отдельной подпапки. Историческая дата первой строки DEV — 09.09.2026; дата в принятом исходном акте №0 — 08.09.2026: принятое №0 не исправлялось.

Текущие реквизиты 9/2026: O400 (`Дата Акз`) = 05.10.2026, C400 договор `06/26 от 03.06.2026`, D400 заказчик `ООО "СТК ВК"`, E400 объект `Архитектурно-строительные решения здания парковки подвижного состава`, F400 шифр `WAA0007144-AC2-КМД`. Состояние DEV не обновлялось по именам папок на диске.

Реальный справочник проекта имеет лист `Марки` и ровно три заголовка: `Марка | Профиль | № АКЗ`; «Краска» отсутствует. Аудит новых форм не доказывает полноту назначений или готовность официальных измерений.

Все шесть новых XLSX найдены. Для каждого проверены сохранённые значения, формулы, объединения, листы, печатные настройки и нормальное открытие в установленном Microsoft Excel 16.0. Native COM: AutomationSecurity=3, EnableEvents=False, UpdateLinks=0, ReadOnly=True, CorruptLoad=0; Close(SaveChanges=False). Save/SaveAs/PrintOut/PrintPreview/экспорт/PDF не вызывались. Источники не копировались в новые акты. Интернет и production не использовались. Число страниц получено через `PageSetup.Pages.Count` при текущих настройках Excel/принтера; визуальный E2E будущих сгенерированных документов в этот аудит не входит.

## 2. NEW TEMPLATE INVENTORY

| № АКЗ | АКЗ filename | Приложение filename | Слоёв по ячейкам | ЛКМ по слоям | Цвет / RAL как записан | Требуемая толщина, мкм | STATUS |
|---|---|---|---|---|---|---|---|
| 1 | АКЗ-1.xlsx | Приложение к АКЗ-1.xlsx | 1 | Norecoat FD Primer Grey | серый; числового RAL нет | 150 | FOUND; CONTENT CONFLICT; PARTIAL |
| 2 | АКЗ-2.xlsx | Приложение к АКЗ-2.xlsx | 2 | Триокор Мастик 4500; Триокор Финиш 5500 | Серый; 7015 (без слова RAL) | 100; `∑1+2=180` — суммарный текст, требование слоя 2 отдельно не задано | FOUND; CONTENT CONFLICT; PARTIAL |
| 3 | АКЗ-3.xlsx | Приложение к АКЗ-3.xlsx | 3 | Триокор Мастик 4500; Триокор Мастик 4500; Триокор Финиш 5500 | Серый; Серый; 7015 (без слова RAL) | 100; 80; 60 | FOUND; CONTENT CONFLICT; PARTIAL |

Суффикс — вариант акта текущего паспорта, а не количество слоёв и не исторический сдвиг. В этих файлах числа случайно совпадают; слойность подтверждена строкой 17 и блоками 18–34.

Exact paths всех форм и их SHA приведены в разделе 13; машинный инвентарь — `selector_inventory.json`.

## 3. №1 STRUCTURE

Лист `4`, видимый, UsedRange по файлу A2:L240. Явная print area отсутствует (пустая и в OOXML, и в Excel); native = 1 страница. Единственный слой: E17=1, E18=`Norecoat FD Primer Grey`, E21=`серый`, E25=150. Значения покрытия можно использовать как исходные требования выбранной формы, но числовой RAL нельзя выдумывать.

Реквизиты: E4=`1/Б2 от `, H4=14.12.2025, E5=`№ 13-25 от 12.08.2025г.`, E6=`ООО «РУССОКО»`, E7 и B44=`2-ЭГ-2019-800/1-Б2-КМД`. Это чужие исторические реквизиты, не сведения текущего Гараж Казахстан. Приложение внутри акта: G8=`=E4`, I8=`=H4` — привязка к собственному старому номеру/дате.

Исторические поля слоя: E19/E20 партии 184.25/140.25, E22 дата `=H4-1`, E23/E24 время 08:00/17:00, E26=`нет`, E27 мокрая толщина 260, E28 сухая 166. Климат E30=`=E22`, E31=15, E32=17.6, E33=35.9, E34=2.8. Это не новые измерения.

Формул всего 5: G8, I8, E22, E30, D49. Volatile/RANDBETWEEN нет. Служебных листов нет. Общая карта подготовки поверхности, приборов и подписей приведена в разделе 7.

## 4. №2 STRUCTURE

Два видимых листа: `4` и `Лист1`. Основной лист A2:M240; print area A1:K51, native 1 страница. `Лист1`: F6:F14, print area отсутствует, native ещё 1 страница; печать всей книги захватит дополнительный лист. F6:F14 содержат старые ЛКМ, партии, цвет, дату-сериал 45999, время и толщины. Формульных ссылок на него в проверенных формулах/defined names нет; это не канонический источник новой партии. Историческое содержимое нельзя случайно выдать в результате.

Номер/дата: E4=`1/2025  от `, H4=11.12.2025. E5=`№ 314-25/1074 от 31 октября 2025 г.`, E6=`ООО «Велесстрой»`, E7 и B44=`ВС.24.ПО-Э1.2-РД-ЭЦ-КМД`. G8=`=E4`, I8=`=H4`.

Первые/вторые номера слоя — E17=1 и H17=2; при этом параметры второго слоя начинаются в I, а не H. Слой 1: E18 Мастик 4500, E21 Серый, E25=100. Слой 2: I18 Финиш 5500, I21=7015, I25=`∑1+2=180`. **Нельзя молча считать I25 толщиной второго слоя 180 или вывести 80 как подтверждённое требование**. Нужна явная трактовка суммарного требования и отдельного слоя.

Партии E19/E20 и I19/I20; даты E22=`=I22-1`, I22=`=H4-1`; время E23/E24 и I23/I24; строка 26 E26=100, I26=256; мокрые E27/I27 вычисляются из сухих E28=115 / I28=220 через общий L27=86. Значение 86 не подтверждает коэффициент для каждого ЛКМ. Климат — E30:E34 и I30:I34. Всего 19 формул на `4`, из них 6 с RANDBETWEEN, включая внепечатные L/M. На `Лист1` формул нет.

## 5. №3 STRUCTURE

Один видимый лист `4`, A2:O240; print area A1:K51, native 1 страница. Три слоя подтверждены E17=1, H17=2, J17=3. Параметры идут в E/H/J: материалы E18/H18/J18; цвет E21/H21/J21; требуемая толщина E25=100, H25=80, J25=60. Сумма явно заданных требований = 240 мкм; это арифметическая проверка, не измерение покрытия.

E4/H4/E5/E6/E7/B44 совпадают с чужими историческими реквизитами №2. G8=`=E4`, I8=`=H4`. Даты нанесения E22=`=H22-1`, H22=`=J22-1`, J22=`=H4-1` — 08/09/10.12.2025 по сохранённым значениям, а не текущие даты. Время E23/E24 и ссылки H23/J23/H24/J24. Партии — строки 19–20 каждого слоя.

Строка 26: E26=100, H26=`∑1+2=180`, J26=`∑1+2+3=240`. Сухие исторические E28/H28/J28 = 115/220/310, мокрые E27/H27/J27 вычисляются через L27=86. Семантику «толщина отдельного слоя» против «накопленная толщина» нельзя устанавливать по возрастающим старым числам. Климат — E/H/J, строки 30–34. Всего 30 формул, из них 7 с RANDBETWEEN, в том числе помощники L/M/N/O вне print area. Служебных листов нет.

## 6. APPLICATION STRUCTURE

Все новые приложения имеют единственный видимый лист `Приложение №1`, заголовок A1:G1, объект A2:G2, двухстрочную шапку A3:G4. A — марка, B — наименование, C — количество, D/E — масса единицы/всех, F/G — площадь единицы/всех. Объединения: A1:G1, A2:G2, A3:A4, B3:B4, C3:C4, D3:E3, F3:G3. Скрытых строк/столбцов нет.

| Вариант | Строки марок | Ёмкость строк ВОМ | Итого | Формулы Итого | Print area | Повтор шапки | Native pages |
|---|---|---|---|---|---|---|---|
| 1 | 5:65 | 61 | 66 | C66=SUM(C5:C65); E66=SUM(E5:E65); G66=SUM(G5:G65) | A1:G66 | 1:4 | 2 |
| 2 | 5:51 | 47 | 52 | C52=SUM(C5:C51); E52=SUM(E5:E51); G52=SUM(G5:G51) | A1:G52 | 1:4 | 1 |
| 3 | 5:98 | 94 | 99 | C99=SUM(C5:C98); E99=SUM(E5:E98); G99=SUM(G5:G98) | A1:G99 | 1:4 | 2 |

В A1 стоят ссылки соответственно на документ о качестве №4/2026 от 24.09.2026, №5/2026 от 25.09.2026, №6/2026 от 30.09.2026. Слова «к акту АКЗ» и нужного суффикса нет. Эти заголовки не совпадают с актами. У всех F/G в теле пустые, G-итог закэширован как 0: **площадь отсутствует, а не доказанно нулевая**. Старые марки, массы, количество и итоги нельзя использовать как новую отгрузку. У №2 уже имеются повторные марки (например К1-13 в строках 6 и 11); будущий импорт сохраняет каждую строку выбранной ВОМ, без агрегации по марке. Формулы только три SUM в каждом приложении, volatile нет.

## 7. FORMULAS / CLIMATE / OLD VALUES

Полная адресная карта всех непустых ячеек и всех формул/кэшей — `cell_dump.txt` и `static_workbooks.json`; значения после нормального открытия Excel — `native_excel.json`. Ошибок Excel в сохранённых ячейках нет; внешних ссылок, connections, QueryTables и VBA-пакетов не обнаружено. Наличие корректного кэша не подтверждает достоверность измерений.

№2 RANDBETWEEN: L31, M31, E32, I32, L33, M33 (6). №3: N31, O31, E32, H32, J32, L33, M33 (7). Других проверенных volatile-функций (RAND/RANDARRAY/NOW/TODAY/OFFSET/INDIRECT/CELL/INFO) нет. В №2 M31/M33 не являются источниками видимого климата, но всё равно остаются случайными формулами в книге. При открытии без сохранения №2 I32 изменился 19.2→20.2; №3 H31 19.6→20.9 и H34 7.42→8.58. SHA на диске остался прежним. Это прямое доказательство недетерминированности старых значений.

Точка росы в №2/№3 рассчитывается как `Tвоздуха - ((100-RH)/5)` в строке 34. Этот алгоритм **не выбран и не подтверждён** пользователем; его нельзя наследовать как утверждённый расчёт. В №1 точка росы — старое число 2.8. Даты нанесения, автоматически вычтенные из H4, также не являются подтверждением реальных работ.

Для всех актов исторические поля подготовки: G11 отметка степени очистки; B13 шероховатость; E13 обезжиривание; I13 обеспыливание; E15 категория подготовки; G37 способ нанесения. Приборы E39:E42 (названия и серийные номера) не доказывают использование/поверку приборов для новых работ. Комиссия E46/E47 — Журавлёв А.Н./Шляхтуров Д.Н.; подписи F49:G49 и I49:K49, дата D49=`=H4`; подписи/должности F50/I50. Факт подписания нельзя переносить. E44 содержит нормативную строку, B43 — безусловное утверждение соответствия; его нельзя оставлять как подтверждение при отсутствии измерений.

Сохраняемые параметры системы покрытия: слойность из 17, ЛКМ из 18, цвет из 21, однозначные требования из 25; условные суммарные значения строки 26 только после определения их смысла. Все партии, дата/время нанесения, измеренные толщины, климат, подготовка, фактическое использование приборов, подписи — отдельные новые данные. Старые E/H/I/J27:28 и общий L27 нельзя использовать для расчёта новых результатов.

Одна запись климата на календарную дату нанесения должна использоваться всеми слоями и актами этой даты. Если слои действительно наносились в разные даты, это должны быть явно подтверждённые даты с соответствующими дневными записями; историческая цепочка H4−1/−2/−3 не подтверждает их. Нельзя генерировать климат по слою. Пока подтверждённого расчёта нет, в TEST — явно отсутствующие данные либо явно маркированный общий TEST-климат по принятому контракту; официальный выпуск STOP.

## 8. TEMPLATE PAIR MATCHING

| № | Файловая пара | Внутри акта | Внутри приложения | Вывод |
|---|---|---|---|---|
| 1 | exact filenames подтверждены | 1/Б2, 14.12.2025 | документ качества 4/2026, 24.09.2026 | FILE MATCH / CONTENT CONFLICT |
| 2 | exact filenames подтверждены | 1/2025, 11.12.2025 | документ качества 5/2026, 25.09.2026 | FILE MATCH / CONTENT CONFLICT |
| 3 | exact filenames подтверждены | 1/2025, 11.12.2025 | документ качества 6/2026, 30.09.2026 | FILE MATCH / CONTENT CONFLICT |

Файловые пары указаны пользователем однозначно, но внутренняя связь АКЗ↔собственное приложение не доказана. Для будущего расширения надо явно утвердить использование этих файлов как форм с перепривязкой реквизитов **только в новых рабочих копиях**. Исходники не исправлять. Договор/заказчик/объект/КМД — из текущей выбранной строки DEV. Акт и его приложение обязаны получить один номер и одну дату DEV: 0→9/2026; 1→9/2026-1; 2→9/2026-2; 3→9/2026-3. Для другого выбранного паспорта меняется базовый номер, а не источник исторических документов.

## 9. DIFFERENCES FROM ACCEPTED №0

Принятые исходники №0: `Акт № 4-2026.xlsx` (лист `4`, один слой Pirocor Protect 057, RAL 7004, 80 мкм) и `1_2026 - Приложение к АКЗ.xlsx` (лист `Приложение АКЗ`, тело 4:20, 17 строк, Итого 21). Их внутренняя пара 1/2026 / 08.09.2026 совпадает. Исходный акт №0 также не имеет явной print area; существующий writer задаёт B2:K50 в TEST-копии. Нельзя приписывать эту output-настройку исходному шаблону.

№1 сохраняет одноколоночную геометрию E, но другое покрытие и чужие реквизиты. №2 — две колонки E/I, отдельный адрес заголовка H17, неоднозначное требование I25, дополнительный видимый лист. №3 — E/H/J и три явных требования. Новые приложения имеют другой лист, A2 объекта, тело с 5, ёмкости 61/47/94 и печать 2/1/2 страницы; writer №0 с A4:G20 и строкой 21 к ним неприменим. №2/№3 требуют очистки всех случайных зависимостей вне печатной области.

Существующие matching/preview/write_act/write_application/validate_saved_pair и GUI сейчас специализируются на №0. PENDING_FILES и generate также содержат точные имена файлов №0; нельзя расширить глобальный allowlist произвольно или ослабить R2 ownership. Реализация №0, принятые R1/R2 ограничения и старые TEST-результаты не менялись. Core SHA совпадает с принятой фиксацией R2: eede991cf2d2ca259f251ef9f79beb49542ea1e536ff2eea6f18cf9c23cd935d.

## 10. MISSING OR AMBIGUOUS FILES

MISSING: NONE. AMBIGUOUS FILES: NONE. Все шесть ожидаемых exact paths найдены; конкурирующих имён в прямом содержимом указанной папки не выявлено. Отдельно имеются CONTENT CONFLICT всех трёх пар, UNRESOLVED требование второго слоя №2 и отсутствие числового RAL у серых слоёв. Это проблемы содержимого, а не отсутствующие файлы. Никакой подстановки из «Другие шаблоны», переименований или перемещений не было.

## 11. IMPLEMENTATION PLAN

Это контракт следующего отдельно разрешённого этапа, не выполненные изменения. Машинная адресная карта — `IMPLEMENTATION_CONTRACT.json`.

1. Сначала закрыть решения по внутренней перепривязке всех трёх пар, требованию второго слоя №2, смыслу строки 26/накопленных измерений, трактовке цветов без явного RAL и политике служебного `Лист1` №2. Если числовой RAL не нужен, сохранить «серый» буквально; не назначать код автоматически. Не выбирать алгоритм точки росы в рамках этого расширения.
2. Оставить ветку №0 и её принятый контракт без изменений. Добавить отдельные описатели вариантов 1/2/3 с exact paths, SHA, ожидаемыми листами, объединениями и адресами из контракта. Маршрутизировать только новые варианты в новые адаптеры, переиспользуя проверенное чтение DEV/ВОМ и файловые проверки. Не менять Stage 1–5, DEV, реальные справочники и исходные формы.
3. Для выбранного текущего паспорта использовать его DEV-реквизиты и `Дата Акз`. Группировать строки выбранной ВОМ по числовому № АКЗ из трёхколоночного справочника. Пустое назначение — отдельное исключение, не №0. Сохранить исходный порядок и каждую повторную строку. Количество/массу/площадь не читать из старого приложения и не вычислять из цвета/профиля. Существующий запрет реального catalog path в DEV-режиме самовольно не снимать; для TEST использовать согласованные изолированные fixtures.
4. Preview каждой новой пары должен показывать вариант, текущий паспорт, номер с суффиксом, DEV-дату, SHA/пути двух форм, список слоёв/требований, provenance каждой строки ВОМ, ёмкость и missing/ambiguous inputs. Никаких автоматических выводов «суффикс=слойность». Для №2 не выдавать preview с выдуманной толщиной второго слоя.
5. В TEST-копиях переписать E4/H4/E5/E6/E7/B44, сохранить связи G8=E4/I8=H4. A1 приложения сформировать как «Приложение № {номер акта} к акту АКЗ № {тот же номер} от {DEV-дата}», A2 — текущий объект. Очистить всё тело A5:Glast и исторические поля всех слоёв, включая helpers вне print area; заменить неподтверждённые факты явными TEST/missing-полями. Не оставлять старое утверждение соответствия как подтверждённое. Назначения комиссии и факт подписания разделить; не копировать подписи как новые.
6. Приложения: N=1..61/47/94; сверх ёмкости — STOP до публикации. Тело фиксировано; лишние строки скрыть в рабочей копии, Итого сохранить видимым, SUM охватывает полное очищенное тело. Не удалять физические строки, не менять объединения и не переносить №0-адреса. Увеличение ёмкости/вставка строк — отдельный scope. Сохранить нужную точность Decimal из текущей ВОМ. Не сжимать большие приложения в одну страницу; проверить штатную пагинацию с повтором 1:4. Для №1 явно зафиксировать output print area после проверки компоновки; у источника её нет.
7. Климат хранить/проверять одной записью по дню, использовать во всех соответствующих колонках всех новых актов. Никаких RANDBETWEEN, случайных помощников и непроверенных вычислений точки росы. Отсутствие свежих измерений оставляет официальный выпуск STOP. TEST-режим не расширяет полномочия официальной выдачи.
8. Новые пары публиковать только после совместной проверки номера/даты, требований, очистки старых фактов/volatile, строк ВОМ, сумм, диапазонов и native Excel открытия. Для каждого нового варианта нужен свой точный allowlist пары и PAIR.json с вариантом/паспортом/SHA/связью с preview. Сохранить durable UUID ownership, проверку идентичности файлов, сохранение первичной ошибки, безопасную уборку только собственных файлов и атомарную публикацию пары из принятого R2. Существующую ветку №0 не заменять общей непроверенной реализацией.
9. Минимальные будущие проверки: новые варианты и их точные SHA/карты; текущая нумерация без исторического offset; повторные строки; пустые/неподдерживаемые назначения; границы ёмкости cap/cap+1; ambiguous №2 STOP; отсутствие старых измерений и случайных формул, в том числе вне печати; общий климат одной даты; pair mismatch; locked app и recovery ownership; native страницы при реальном N. Регрессия №0 — существующие принятые R1/R2 проверки без пересоздания старых TEST-артефактов. User E2E — только после нового implementation и отдельного предъявления новых результатов.

## 12. CHANGED FILES

Только новые audit evidence в текущем каталоге. Product/source/templates/DEV/catalog/старые TEST: изменений 0. Созданы audit.py, prepare_report.py, selector_inventory.json, sources_before.json, static_workbooks.json, cell_dump.txt, native_excel.json, sources_after.json, integrity.json, IMPLEMENTATION_CONTRACT.json, AUDIT_REPORT.md, AUDIT_MANIFEST.json. Последний манифест содержит SHA всех остальных evidence. В ходе сбора ограничена область хэширования до релевантных 534 файлов; технические node_modules/.codebase-index/__pycache__ исключены. Промежуточные процессы сборщика прекращены до открытия Excel; источники ими не сохранялись. Чужой открытый Stage6 GUI не закрывался и не использовался.

## 13. SOURCE SHA BEFORE / AFTER

Контролировались 534 файла: принятый guard предыдущей фиксации, текущие исходники Stage 1–6/тесты/DEV, старые TEST-артефакты, прямые файлы первого паспорта, реальный catalog и выбранная ранее ВОМ. По окончании native inspection и повторно перед финальным манифестом `changes=[]`. Полный exact-path реестр: sources_before.json и sources_after.json.
''')
add('### Exact paths и SHA новых форм')
for x in inventory['files']:
    p=x['path']; nb=next(b for b in native if b['path']==p)
    add(f"- `{p}`\n  - BEFORE: `{before[p]}`\n  - AFTER: `{nb['sha_after']}` — MATCH\n  - Листы / native pages / print: "+'; '.join(f"{s['name']} / {s['pages']} / {s['print_area'] or '(не задана)'}" for s in nb['sheets']))
add('### Принятые №0, DEV и другие защищённые входы')
checks=[folder/'Акт № 4-2026.xlsx',folder/'1_2026 - Приложение к АКЗ.xlsx',ROOT/'DEV/ВЗМК_DEV.xlsx',folder.parent/'Справочник марок.xlsx',ROOT/'akz_stage6/core.py',ROOT/'akz_stage6/gui.py',ROOT/'validation/stage6_fix_r1_r2_20261009_094707/NEW_TEST_9_2026/ТЕСТ АКЗ №0.xlsx',ROOT/'validation/stage6_fix_r1_r2_20261009_094707/NEW_TEST_9_2026/ТЕСТ Приложение к АКЗ №0.xlsx']
for p in checks:add(f'- `{p}`\n  - BEFORE = AFTER: `{before[str(p)]}`')
add('''## 14. READY TO IMPLEMENT АКЗ-1/-2/-3

**PARTIAL** для всех трёх вариантов. Формы существуют, нативно открываются, структура и фиксированные ёмкости установлены; карта расширения подготовлена. Автоматический перенос без разрешения конфликтов содержимого не готов.

- №1: явные один слой/150 мкм; требуется перепривязка чужих реквизитов и приложения, решение о буквальном «серый» без числового RAL и явной print area в будущей копии.
- №2: дополнительно блокирует неоднозначное требование второго слоя; требуется политика `Лист1`, исключение случайного климата и подтверждение семантики толщин.
- №3: требования 100/80/60 явные; требуются перепривязка, исключение случайного климата и исторических измерений, подтверждение семантики накопленных толщин.

Официальная выдача: **NO / STOP**, свежие измерения и подтверждённый климатический расчёт этим аудитом не получены. Никакое новое разрешение на official mode не подразумевается. Implementation и Stage 7 не начаты. После отчёта работа остановлена.
''')
(OUT/'AUDIT_REPORT.md').write_text('\n'.join(report),encoding='utf-8')
after={p:hashlib.sha256(Path(p).read_bytes()).hexdigest() if Path(p).is_file() else 'MISSING' for p in before}
changes=[{'path':p,'before':h,'after':after[p]} for p,h in before.items() if h!=after[p]]
write('sources_after.json',after)
integrity=read('integrity.json');integrity.update(final_recheck=True,changes=changes)
write('integrity.json',integrity)
assert not changes,changes
manifest={'kind':'READ_ONLY_AUDIT_NOT_IMPLEMENTATION','ready':'PARTIAL','guarded_files':len(before),'source_changes':changes,'native_files':len(native),'evidence':{p.name:hashlib.sha256(p.read_bytes()).hexdigest() for p in sorted(OUT.iterdir()) if p.is_file() and p.name!='AUDIT_MANIFEST.json'}}
write('AUDIT_MANIFEST.json',manifest)
print(json.dumps({'report':str(OUT/'AUDIT_REPORT.md'),'guarded':len(before),'changes':changes,'evidence':len(manifest['evidence']),'ready':'PARTIAL'},ensure_ascii=False))
