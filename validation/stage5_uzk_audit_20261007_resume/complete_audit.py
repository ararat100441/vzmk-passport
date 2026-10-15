import json,csv,re,hashlib
from pathlib import Path
from collections import Counter,defaultdict
from datetime import datetime,timezone,timedelta
import openpyxl
OUT=Path(__file__).resolve().parent;APP=OUT.parent.parent;ROOT=APP.parent
def read(p):return json.loads(Path(p).read_text(encoding='utf-8-sig'))
def save(n,d):(OUT/n).write_text(json.dumps(d,ensure_ascii=False,indent=2,default=str),encoding='utf-8')
def sha(p):return hashlib.sha256(Path(p).read_bytes()).hexdigest()
def table(headers,rows):
    def cell(x):return str(x).replace('|','/').replace('\n',' / ') if x is not None else 'EMPTY'
    return '\n'.join(['| '+' | '.join(headers)+' |','| '+' | '.join('---' for _ in headers)+' |']+['| '+' | '.join(cell(v) for v in r)+' |' for r in rows])
choices={'project':'Гараж Казахстан','mark_catalog':'WORKING_CATALOG_ABSENT; user explicitly chose to record absence; E2E fixture not used',
         'vom_path':str(ROOT/'Гараж казахстан'/'Паспорт'/'9 от 06.10.2026'/'ВОМ отгрузочный счет 183  (5).xls'),
         'vom_row_scope':'ONLY_VISIBLE_ROWS_CURRENT_SHIPMENT','implementation':'NOT_AUTHORIZED'}
save('user_choices.json',choices)
compat=read(OUT/'compatibility.json');compat['scope_status']='USER_CONFIRMED_VISIBLE_ONLY';compat['current_Q']=33
save('compatibility.json',compat)
current=[r for r in compat['rows'] if r['Q_visible_only']>0]
save('current_passport_uzk.json',{'project':choices['project'],'passport':'9/2026','scope':choices['vom_row_scope'],
                              'rows':current,'controlled_quantity':33,'automatic_geometry':'BLOCKED_MARK_CATALOG_MISSING','document_created':False})
book=openpyxl.load_workbook(APP/'DEV'/'ВЗМК_DEV.xlsx',data_only=True);s=book['Паспорт качества']
firsts=defaultdict(list)
for row in s.iter_rows(min_row=2):
    vals=[c.value for c in row]
    if isinstance(vals[1],str) and re.match(r'^1/',vals[1]):
        firsts[vals[0]].append({'row':row[0].row,'number':vals[1],'object':vals[4],'code':vals[5]})
book.close()
inventory=read(APP/'validation'/'stage5_uzk_audit_20261007_stop'/'file_inventory.json')
classified=[]
for old in inventory:
    if not old['classification'].startswith('first-passport'):continue
    text=old.get('content_excerpt','');act=re.search(r'Акт\s*№\s*(.*?)\s+от\s+(\d{2}\.\d{2}\.\d{4})',text)
    number=act[1] if act else None
    matches=[name for name,rs in firsts.items() if len(rs)==1 and rs[0]['number']==number and str(rs[0]['object'] or '') in text and str(rs[0]['code'] or '') in text]
    classified.append({'path':old['path'],'sha256_snapshot':old['sha256'],'size_bytes':old['size_bytes'],
                       'modified_snapshot':old['modified_local_utc_plus_04'],'act_snapshot':number,'date_snapshot':act[2] if act else None,
                       'DEV_exact_number_object_code_matches':matches,'classification':'FIRST_PASSPORT_CONFIRMED_IN_PRIOR_SNAPSHOT' if len(matches)==1 else 'CANDIDATE_NEEDS_EXPLICIT_CONTEXT',
                       'selected':False,'live_content_reopened':False})
selected=ROOT/'Гараж казахстан'/'Паспорт'/'1_2026 - УЗК.docx'
for r in classified:
    if r['path']==str(selected):r.update(classification='SELECTED_PROJECT_FIRST_PASSPORT',selected=True,live_content_reopened=True)
save('first_passport_classification.json',classified)
with (OUT/'first_passport_classification.csv').open('w',encoding='utf-8-sig',newline='') as f:
    writer=csv.DictWriter(f,fieldnames=list(classified[0]));writer.writeheader();writer.writerows(classified)
eligible=[p for p in selected.parent.iterdir() if p.is_file() and re.match(r'^1_2026\s*-\s*УЗК\s*\.docx$',p.name,re.I)]
for d in selected.parent.iterdir():
    if d.is_dir() and re.match(r'^1\s+(?:от\s+)?\d{2}\.\d{2}\.2026',d.name):
        eligible.extend(p for p in d.iterdir() if p.is_file() and re.match(r'^1_2026\s*-\s*УЗК\s*\.docx$',p.name,re.I))
assert eligible==[selected],'STOP selected first passport template ambiguous'
save('selected_template_policy.json',{'project':'Гараж Казахстан','canonical_first_number':'1/2026','passport_root':str(selected.parent),
                                    'eligible_candidates':[str(p) for p in eligible],'selected':str(selected),
                                    'fallback_checked':str(ROOT/'Шаблоны'/'УЗК.docx'),'fallback_exists':(ROOT/'Шаблоны'/'УЗК.docx').is_file(),
                                    'recursive_fuzzy_selection':False})
before=read(OUT/'preservation_before.json')
for r in read(OUT/'word_inputs.json'):before[r['path']]=r['sha256']
v=read(OUT/'native_vom.json');before[v['path']]=v['before_sha']
save('preservation_before.json',before)
after={p:sha(p) for p in before}
assert before==after,'STOP source drift during audit'
save('preservation_after.json',after)
native=read(OUT/'native_word.json');words=read(OUT/'word_structure.json');samples=read(OUT/'registry_samples.json')
checkpoint=read(APP/'DEV'/'DEV_UZK_CHECKPOINT.json');prep=read(OUT/'prepare.json');counts=compat['counts']
scope_counts=Counter(r['classification'] for r in classified)
macros=read(OUT/'vba_procedures.json')
summary={'status':'AUDIT_COMPLETED_WITH_MISSING_REQUIRED_INPUT','ready_to_implement':False,'user_choices':choices,
         'registry_counts':counts,'first_passport_classification_counts':dict(scope_counts),
         'protected_count':len(before),'protected_changes':[],'accepted_baseline_checks':22,
         'source_sha_before':checkpoint['source_sha256'],'source_sha_after':after[checkpoint['source_path']],
         'dev_sha_initial':checkpoint['dev_initial_sha256'],'dev_sha_after':after[checkpoint['dev_path']],
         'native_excel_readonly':read(OUT/'native_excel.json')['status'],'native_word_readonly':[r['status'] for r in native],
         'word_visual_pages_reviewed':6,'new_UZK_docx_created':False,'VBA_executed':False,'decrement':False,
         'production_VZMK_read_or_hashed':False,'network_access':False,
         'blocker':'Working project mark catalogue absent; 4 current / 19 garage registry profiles cannot be established; 6 other-project rows outside selected context.'}
save('audit_summary.json',summary)
sample_rows=read(OUT/'registry_structure.json')['sheets'][0]['raw_rows'][1:]
rawtable=table(['ROW','A: Заказ','B: Марка элемента','C: raw quantity','Remaining by approved rule'],[[r['row'],r['values'].get('A'),r['values'].get('B'),r['values'].get('C'),r['values'].get('C',1)] for r in sample_rows])
currenttable=table(['УЗК row','Марка','ВОМ cell','ВОМ qty','Registry raw','Remaining','Q','Profile status'],[[r['row'],r['mark'],'D'+str(r['vom_matches'][0]['row']),r['vom_matches'][0]['quantity'],r['raw_quantity'],r['remaining'],r['Q_visible_only'],'MARK_CATALOG_MISSING'] for r in current])
fullmatch=table(['Registry row','Марка','ВОМ row','ВОМ qty','Hidden','Q all','Q current visible'],[[r['row'],r['mark'],r['vom_matches'][0]['row'],r['vom_matches'][0]['quantity'],r['vom_matches'][0]['hidden'],r['Q_all_rows'],r['Q_visible_only']] for r in compat['rows']])
metadata=[]
for p in [Path(checkpoint['source_path']),Path(checkpoint['dev_path'])]+[Path(x['path']) for x in read(OUT/'word_inputs.json')]+[Path(v['path'])]:
    st=p.stat();metadata.append({'path':str(p),'size_bytes':st.st_size,'modified_utc_plus04':datetime.fromtimestamp(st.st_mtime,timezone(timedelta(hours=4))).isoformat(),'sha256':sha(p)})
save('relevant_file_inventory.json',metadata)
metatable=table(['EXACT PATH','SIZE BYTES','MODIFIED UTC+04','SHA-256'],[[r['path'],r['size_bytes'],r['modified_utc_plus04'],r['sha256']] for r in metadata])
mapping=[
['Body P1–P3','ООО / Волжский завод металлоконструкций / ЛНК ООО ВЗМК','project first-passport template','CONSTANT'],
['Body P5: акт','1/2026','DEV Паспорт качества B353; current B400=9/2026','DYNAMIC'],
['Body P5: дата','04.09.2026','DEV N353 УЗК Дата; current N400=01.10.2026','DYNAMIC'],
['T1 R1 C2','ООО «ВЗМК»','template; not DEV D Заказчик (ООО СТК ВК)','CONSTANT'],
['T1 R2 C2: объект','Архитектурно-строительные решения здания парковки подвижного состава','DEV E current row','DYNAMIC by project'],
['T1 R2 C2: шифр','WAA0007144-AC2-КМД','DEV F current row; parenthesized after object','DYNAMIC by project'],
['T1 R3 C1','А1212 MASTER №3121397; КОУ-2; СО-3 №6131; СО-2 №6153; поверки','project template; unchanged in relevant examples','CONSTANT'],
['T1 R6–R16 C1','К1-5 … К1-34 (header № п/п)','visible current VOM B; display original mark','DYNAMIC'],
['T1 data C2: L','810 / 1794; 270 / 598','Stage 4 geometry × Q; per-mark groups','DYNAMIC'],
['T1 data C2: S','9 / 12; 9 / 14','Stage 4 geometry; not multiplied','DYNAMIC'],
['T1 data C3','Иванов С.В. H630; Евсеев А.В. 2Т33; Фомин С.А. 2ХВ4','project template; same in compared examples','TEMPLATE CONSTANT'],
['T1 data C4','04.09.2026','same DEV N date as act header','DYNAMIC'],
['T1 data C5','5 / 2,5','frequency_angle(S), aligned with C2 lines','DYNAMIC'],
['T1 data C6','70 / 65','frequency_angle(S), aligned with C2 lines','DYNAMIC'],
['T1 data C7','ДНО','project template; same compared examples','TEMPLATE CONSTANT'],
['T1 data C8','годен','project template; same compared examples','TEMPLATE CONSTANT'],
['T2 R1 C1–C3','Контроль выполнил / underline / Журавлев А.Н.','project template','CONSTANT'],
['T2 R2 C1','II; удостоверение №0071-0584; 01.08.2025; НОАП АНОО ДПО «ЦПиПК»','project template','CONSTANT'],
['Headers / footers / PAGE','Absent','no PAGE fields or header/footer parts; mark hyperlinks separately documented','NO DYNAMIC FIELD'],
['T1 data C1 hyperlink','11 historical PDF links, e.g. К1-5.pdf','No approved current drawing-link source; replace old hyperlink with plain current mark unless separately approved','DYNAMIC / REMOVE STALE LINK'],
['Profile / quantity separate columns','Absent','lookup and Q are intermediate inputs, represented through L','DO NOT ADD COLUMNS']]
save('document_field_mapping.json',[dict(zip(['location','current_value','source','mode'],r)) for r in mapping])
fieldtable=table(['DOCUMENT LOCATION','CURRENT VALUE','SOURCE','DYNAMIC / CONSTANT'],mapping)
report=f'''# Stage 5 предварительный аудит УЗК после уточнения source path

Дата 07.10.2026, Europe/Samara. Проект выбран пользователем: Гараж Казахстан. Текущий ВОМ выбран пользователем: паспорт 9 от 06.10.2026, только видимые строки текущей отгрузки. Это аудит и предложенный контракт, implementation отсутствует.

## 1. ACTUAL STATE

Stage 1–4 accepted, Stage 4 User E2E закрыт — текущая авторизация пользователя; повторная приёмка не запускалась. Все 22 accepted hashes из предыдущего audit checkpoint подтверждены снова. DEV ВЗМК_DEV.xlsx: 8e56f68518c70cade9efb09084dd67cd5d72e07e3fdc813b8515030e9cb98e80. Accepted Профили.xlsx: 39f14b040d3a97365f9e44f5fdc01c87e1f2ce068d5d4f26458a91602759ca91. Sidecars и Stage 1–4 code сохранены.

{len(before)} защищённых файлов BEFORE=AFTER; список в preservation_before/after.json. Original УЗК использовался для первичного копирования и fingerprints; содержимое исследовано только из DEV. Production ВЗМК.xlsx не открывался и не хешировался. Предыдущие audit/evidence проверены по manifest и не изменялись.

## 2. FILE INVENTORY

{metatable}

Authoritative original — production source, read-only fingerprint/copy only. DEV — единственный development input. Первый DOCX гаража — выбранный project first-passport template; второй — result/comparison example; ТЭМПО 2 — другой project first-passport example для сравнения нумерации, constants и таблицы; выбранный ВОМ — source текущего паспорта 9.

Все 339 DOCX повторно не инвентаризировались. Повторно прочитаны только 3 relevant DOCX. Классификация 43 first-passport candidates по предыдущим content excerpts и DEV exact number/object/code сохранена в first_passport_classification.csv/json: {dict(scope_counts)}. Для остальных проектов это snapshot classification, не live template selection. Для гаража live проверен единственный eligible 1_2026 - УЗК.docx в корне паспорта/канонической first-passport folder; recursive fuzzy selection отсутствует. Два экземпляра первого акта Щёкиназот в старой inventory имеют одинаковые bytes, но не выбирались как один автоматически.

Точный проверенный fallback Документы\\Шаблоны\\УЗК.docx отсутствует. Другие предполагаемые внешние locations не искались; при выбранном проектном шаблоне fallback не требуется.

## 3. УЗК.XLSM STRUCTURE

SHEET: УЗК. State: visible; hidden/veryHidden листов нет. Used range A1:C26; ROWS=26 включая header; COLUMNS=3; 25 рабочих строк. HEADERS: A1 «Заказ»; B1 «Марка \\nэлемента»; C1 «Кол-во (шт)». PURPOSE: список заказов/марок и заданного пользователем remaining количества для УЗК.

Excel Tables=0; formulas=0; merged cells=0; workbook/local named ranges=0; скрытых строк/столбцов=0. Workbook structure/windows protection=false; sheet protection=false. External workbook links/relationships/connections=0. AutoFilter A1:C21 — не охватывает рабочие строки 22–26. Поэтому registry читать до последней фактической строки 26, без ограничения filter range.

Validation: list только A9 и A26, allowBlank=true, «Гараж Казахстан,Гранит», ошибка «Выберите заказ из списка для этой марки». Это подтверждает необходимость composite key заказ+марка. C не имеет quantity-validation или default formula.

VBA present: xl/vbaProject.bin, 414720 bytes, SHA 7464ff05e1973954114ad336ff10464fa34e87dd3ec5ae4a5f4a76d9913a2dcf. Native safe static read доступен, protection=0. ЭтаКнига: 31 line; Module1: 2404 lines; Лист8: 0. Извлечён {len(macros)} procedure declaration, полный список vba_procedures.json; среди public — UpdateAll, GenerateOrderReport, DeleteOrderFromButton. Event Workbook_SheetChange присутствует; Workbook_Open declaration не обнаружена. References: VBA 4.2, Excel 1.9, stdole 2.0, Office 2.8; broken=false.

Код унаследован от большой производственной книги: обращается к отсутствующим здесь Сводная/Сборка/АКЗ/ОГЗ/Отгрузка. FillOneOrderOnlySheet читает B как mark и A как order, создаёт validation; правила blank quantity=1 и passport-close decrement этим кодом не доказаны. DeleteOrderData содержит удаление строк заказов, что не является согласованной будущей операцией passport-close. Ничего из VBA не исполнялось.

Native Excel 16.0: ReadOnly=true, AutomationSecurity=3, EnableEvents=false, UpdateLinks=0, normal open без repair, Close(false), SHA unchanged. Первичная ошибка PowerShell binding optional parameters сохранена отдельно; исправлен только audit harness, повторный normal open успешен, repair не применялся.

## 4. REGISTRY BUSINESS MAPPING

| FIELD | COLUMN / HEADER | EXAMPLE | MEANING | CONFIDENCE / EVIDENCE |
| --- | --- | --- | --- | --- |
| Project/order | A / Заказ | A8=Гараж Казахстан; A2=2 ТЭМПО | Проектный scope, часть ключа | HIGH: raw rows + static FillOneOrderOnlySheet |
| Mark | B / Марка элемента | B8=Бс1-7; B14=К1-24 | Отправочная марка, не профиль и не порядковый номер | HIGH: VOM matching + static VBA |
| Quantity | C / Кол-во (шт) | C8=EMPTY; C10=4; C24=15 | Remaining: blank=1; numeric=remaining | HIGH under approved user rule; workbook alone does not prove remaining semantics |
| Profile | отсутствует | — | Нужен проектный справочник | HIGH structural evidence |
| Passport/date/other number-ID | отсутствуют | — | Реквизиты брать из DEV ВЗМК registry, не из row index | HIGH structural evidence |

Количество/остаток — один C, отдельного столбца «номер/остаток» нет. Номера строк Excel не являются business ID. Current A9/A26=Гараж Казахстан; список validation не означает две записи или автоматическое совпадение с Гранит.

## 5. REGISTRY SAMPLE EVIDENCE

Все 25 raw rows, так как малый реестр:

{rawtable}

Детерминированно: first=row2; middle=row14 (13-я из 25); last=row26. Первые 5 blank rows: 2,3,5,7,8. Первые 5 numeric: 4,6,10,11,14. Blank=16; numeric=9, все положительные целые, диапазон 1–15. По approved rule суммарный remaining=62; это не количество текущего паспорта. Duplicate raw/normalized composite keys=0; repeated marks внутри текущего реестра не обнаружены. Нулей, отрицательных, дробных, text quantities и formula quantities нет.

Special: A9/A26 list validation с двумя заказами; A22:C26 за пределами filter; два разных написания order context ТЭМПО 2 в DEV и 2 ТЭМПО в УЗК требуют явного mapping, но этот проект сейчас не выбран.

blank=1 применено как утверждённая user semantics. В файле пустота действительно EMPTY, не формула и не 0; собственного defaults/decrement mechanism для C не найдено. Историческое списание не моделировалось.

## 6. УЗК.DOCX STRUCTURE

Selected: Гараж казахстан\\Паспорт\\1_2026 - УЗК.docx, hash в разделе 2. Word native: 2 pages, 1 section, A4 landscape, margins 850 twips (~15 мм), 155 native paragraphs, 2 tables. OOXML w:p elements=137; это отдельная structural metric. Body direct paragraphs=8.

Table1: 16 physical rows, 8 logical columns в control rows, 9 underlying grid columns; rows1–4 реквизиты/оборудование/blank spacer, row5 control header, rows6–16=11 mark rows. Horizontal spans: first two rows 2+7, equipment/spacer span9, control rows 1+2+1+1+1+1+1+1. Vertical merges=0. Table2: 2 rows, 3 grid columns, second row span3 (signature/qualification). AutoFit=false. Repeating headers=0. Totals rows=0.

Header/footer parts=0, PAGE fields=0, sections page-number fields отсутствуют; content controls=0; shapes/inline shapes/media/textboxes=0; protection=-1, revisions=0, TrackRevisions=false. Подпись — текст Журавлев А.Н. и underline, не вставленное изображение; stamps/images отсутствуют.

Fixed text: фирма/ЛНК, заказчик ООО ВЗМК, оборудование, поверки, сварщики/клейма, ДНО/годен, оператор и удостоверение. Номерного нормативного документа контроля в тексте нет; есть header «согласно ТУ». Профиль, отдельное количество, договор, сталь не являются отдельными полями этой формы.

Word native read-only normal opens трёх DOCX успешны, macros disabled, no repair/no save; все SHA unchanged. Шесть страниц просмотрены как native EMF→PNG без PDF. Existing layout splits mark К1-12 между страницами и не повторяет table header; во втором акте аналогично разделена К1-19. На последней строке первого акта частота/угол визуально тесно расположены. Источники не исправлялись.

В выбранном первом DOCX 11 mark hyperlinks: native Word Fields.Count=11, все type88 HYPERLINK; это relationships w:hyperlink, поэтому XML instrText/fldChar=0 не означает отсутствие hyperlink fields. Ссылки ведут к старым локальным PDF сборок (например К1-5.pdf); они не открывались. Во втором гаражном и ТЭМПО examples native Fields.Count=0. PAGE fields отсутствуют. Evidence: word_hyperlinks.json и native_word_fields.json.

## 7. DOCUMENT FIELD MAPPING

Локации C — physical/logical cell order control row, не underlying grid start. Полный mapping:

{fieldtable}

Дата доказана отдельным N «УЗК Дата» DEV registry: N353=04.09.2026 при I353=09.09.2026; N360=10.09.2026 при I360=15.09.2026; N400=01.10.2026 при I400=06.10.2026. Formula N=Таблица1[[#This Row],[Дата окончания ]]-5. M «ВИК Дата»=end−7 и отличается; не использовать M, today или date папки.

Номер выбранного проекта: первый акт 1/2026=DEV B353, второй 2/2026=DEV B360, текущий паспорт B400=9/2026. Для гаража нет доказанного дополнительного УЗК suffix; предложенный номер текущего акта=9/2026. Нумерация страниц отсутствует. ТЭМПО example=1/2/2026 — отдельный проектный формат, не заменять общей схемой гаража.

CONSTANT означает сохраняемый текст проектного шаблона в proposed generation contract, а не доказательство текущих кадров/оборудования/результата будущего контроля.

## 8. MARK ROW MODEL

Одна physical control row = одна марка. В C2 одной строки несколько L×S, разделённых w:br; C5 frequency и C6 angle содержат aligned line lists того же порядка. Сварщики C3 остаются один блок на марку; C1 mark, C4 date, C7/C8 result повторяются только для разных марок, не для каждой толщины.

Для одной марки несколько толщин не требуют нескольких Word rows. После внутри-марочного объединения одинаковой S число lines может стать 1. Между марками не объединять.

Первый garage template имеет 11 mark rows, второй result=6, ТЭМПО example=7: длина control region реально меняется. Нет защиты/content controls, препятствующих изменению; технически возможно клонировать первую control row и заменить только data rows. Native add/remove/generation не запускались. Proposed: preserve реквизиты rows1–5 и отдельную Table2, создать ровно число нужных марок, заменить всю старую geometry/frequency/angle data, не переносить результаты предыдущих марок. Page count после будущего изменения не фиксировать равным 2; потребуется native validation без изменения originals.

## 9. TEMPLATE COMPARISON

Project vs exact global: NOT_RUN, проверенный global fallback отсутствует. Это не повод выбирать другой project template.

Relevant comparison: garage first/second и ТЭМПО first имеют A4 landscape, 2 native pages, 2 tables, 8 logical control columns, одинаковые оборудование/поверки/сварщиков/оператора/удостоверение/ДНО/годен. Различаются object/code, numbering, dates, widths/row layout и marks/geometry. Garage table1 rows=16/11; Tempo=12; mark rows=11/6/7.

В ТЭМПО исторические строки К1-3,К1-5,К1-15,К1-22 содержат geometry 8 мм затем 12 мм, но frequency 2.5 затем 5 и angle65 затем70 — reverse относительно approved rule: 8 несовпадающих segments в 4 rows. Это доказательство необходимости пересчитывать dynamic cells по S, не копировать их значения из образца. Исторические документы не исправлялись; никаких новых правил из них не выводилось.

## 10. APPROVED STAGE 5 CALCULATION CONTRACT

Registry key=(selected project,normalized mark), C blank→1, numeric→remaining, Q=min(current visible VOM quantity,remaining). Source spelling сохранять; normalization только matching; collisions STOP. Scan registry A2:C26, не AutoFilter A1:C21. Source quantity допускает только доказанные безопасные значения; текущие C положительные целые. Не путать отсутствующую mark с отсутствующим/невалидным registry.

Project catalog exact columns Марка / Профиль / Краска / № АКЗ; Stage5 использует первые два. E2E catalog не является рабочим source. Lookup только accepted local catalog with FOUND/FOUND_PREFERRED; AMBIGUOUS/CONFLICT/NOT_FOUND блокируют automatic generation.

Only approved ДВУТАВР_ПРОКАТНЫЙ_1: wall L=h−2t,S=s; flanges L=2b,S=t. L_total=L_single×Q, S unchanged. Same S summed внутри марки; разные марки всегда отдельно. S<12 →5 MHz/70°, S>=12→2.5/65. Read-only accepted local probe: 35Б1 FOUND_PREFERRED; 328×6+348×9, Q3→984×6+1044×9; 11.5→5/70,12→2.5/65 PASS. Это пример accepted catalog, не profile matching реальных текущих марок.

Другой неподтверждённый тип → UZK_RULE_NOT_CONFIGURED, automatic generation STOP. Внутренний Stage4 status «НЕ НАСТРОЕНО» не менять; proposed Stage5 adapter только отображает согласованный статус.

## 11. REAL DATA COMPATIBILITY

Выбранный XLS ВОМ: native file format56, used A1:K507, 501 mark rows B4:B504, quantities D, name C; active AutoFilter и FilterMode=true. Только104 видимые mark rows относятся к выбранной текущей отгрузке по прямому подтверждению пользователя. Их суммарная масса=15822.14 кг, совпадает с DEV G400. Visible quantity=288. Полная таблица даёт 2191 изделия/156879.26 кг; footer D505=2453 и I505=172595.01 — не использовать как total текущей видимой отгрузки. Эти значения не исправлялись.

Все19 registry marks гаража встречаются в полном ВОМ, из них4 видимы и15 скрыты. Наивное чтение всех rows дало бы Q54; approved current visible scope даёт Q33:

{currenttable}

Полное сопоставление19 registry rows:

{fullmatch}

Рабочий exact Гараж казахстан\\Справочник марок.xlsx отсутствует; пользователь явно выбрал зафиксировать отсутствие. Найденный E2E файл не использовался. Поэтому current4 marks/Q33 заблокированы до profile mapping. Для всех19 garage registry rows профиль не установлен; 6 rows «2 ТЭМПО» вне выбранного project scope. Real registry verified-supported=0, unresolved-profile=25; это не означает25 unsupported. Actual unsupported/AMBIGUOUS/CONFLICT/NOT_FOUND counts UNKNOWN: real profile lookup NOT_RUN из-за отсутствующих профилей. Названия «Балка»/«Прогон стеновой» и габариты E/F/G ВОМ не заменяют точный профиль или НД. Не угадывать по ним двутавр/трубу/другую family.

Compatibility gap принятого Stage1 reader: он принимает .xlsx, не выбранный .xls; read_vom читает rows без учета hidden/filter. Stage1–4 не менялись. Для Stage5 нужен отдельный согласованный read-only input adapter/snapshot для visible selection, либо отдельно подготовленный пользователем корректный current XLSX. Автоматическая конверсия здесь не выполнялась.

## 12. NO-UZK BEHAVIOR

Если после успешной валидации всех inputs/matching Q>0 ни у одной марки, УЗК.docx не создаётся; будущий Документ о качестве не ссылается на УЗК. Текущий результат не NO-UZK: есть4 required marks/Q33. Отсутствующий catalog означает BLOCKED, не отсутствие контроля.

## 13. REGISTRY MUTATION BOUNDARY

В audit registry не изменялся. В обычной будущей generation decrement/delete тоже запрещён. Explicit passport close/decrement — отдельная будущая безопасная операция; inherited VBA DeleteOrderData не использовать как готовое close API. Source original только fingerprint; Stage5 tests/development только УЗК_DEV.xlsm, без automatic sync/refresh.

## 14. USER DECISIONS REQUIRED

Уже решено: проект Гараж Казахстан; exact ВОМ паспорта9; только visible rows; отсутствие рабочего справочника фиксировать; E2E не подменять. Date, numbering выбранного проекта и row model установлены по файлам, повторных вопросов по ним нет.

USER DECISION REQUIRED1 — перед implementation определить authoritative рабочий источник Марка→Профиль для гаража. Текущее решение пользователя — оставить отсутствие зафиксированным, поэтому это input blocker, не просьба снова подтвердить уже выбранный вариант. Evidence: exact рабочий catalog отсутствует;4 current marks и19 registry marks без profile. Options для отдельной следующей задачи: предоставить рабочий exact catalog; отдельно разрешить его создание/наполнение проверенными данными; оставить automatic UZK generation заблокированной. В этом audit каталог не создаётся.

Не выбранный ТЭМПО не требует решения сейчас. Если Stage5 должен обслуживать его, потребуется explicit 2 ТЭМПО↔ТЭМПО2/project folder/catalog/number format mapping, не выводить этот alias из перестановки слов автоматически.

## 15. PROPOSED STAGE 5 IMPLEMENTATION CONTRACT

A Inputs: selected project/current passport + checkpointed DEV ВЗМК registry, УЗК_DEV registry, working project mark catalog, accepted local profiles/sidecars, selected exact first-passport template. Для этого audit user selected passport9 current, не автоматически next10.

B Registry mapping: sheet УЗК, A project/B mark/C remaining, header row1, allactual rows2–26; no filters as data boundary; composite key and collision STOP; blank=1 and positive integer quantities; zero future=0, never blank-default it.

C Matching: exact approved project mapping, safe mark normalization without changing display, working catalog Марка→Профиль, accepted Stage4 resolved lookup only. Missing inputs/alias/collisions stop before document creation.

D Quantity: current visible VOM rows only, B mark/D quantity for selected XLS; freeze selected scope and source hash. Q=min(qty,remaining). No grouping different marks. No use of hidden rows or stale full-table footer as current totals. Proposed input adapter must live in Stage5; no Stage1–4 edits implied.

E Geometry: use accepted Stage4 one-item formula and statuses; multiply only L, combine equal S within mark, compute frequency/angle from S. Unsupported blocks whole automatic document.

F Word model: one mark row, C2/C5/C6 aligned segment lists; preserve top5 rows and Table2; replace all historical data rows. No new profile/quantity columns. Historical mark hyperlink targets must not survive row cloning; use plain current mark until an authoritative drawing-link source is separately approved. Native dynamic generation and pagination gates will be required in later implementation, not performed now.

G Header: number DEV B, UZK date DEV N, object E plus code F in parentheses. Preserve project template company/customer/equipment/personnel/certification/result constants. Verify expected template old values before mutation of a future disposable output. Do not replace ОООВЗМК customer with DEV external customer.

H No-UZK: only after successful input validation and zero required matches; no file and no future quality-document reference.

I STOP: checksum/checkpoint drift, missing/ambiguous template/catalog/registry/project keys, invalid quantities, duplicate normalized marks, unsafe profile lookup, unsupported UZK rule, recovery/repair, required VBA execution or production write. Current blocker: missing working mark catalog.

J Boundary: source read-only; outputs later only authorized destinations; generation read-only registry; close/decrement separate explicit operation. Exact current DEV snapshot remains fixed until separate user refresh decision.

## 16. DEV UZK CHECKPOINT

Source path: {checkpoint['source_path']}. DEV path: {checkpoint['dev_path']}. Source BEFORE copy={checkpoint['source_sha256']}; DEV initial={checkpoint['dev_initial_sha256']}; source AFTER copy={checkpoint['source_sha256_after_copy']}; source AFTER audit={after[checkpoint['source_path']]}; DEV AFTER audit={after[checkpoint['dev_path']]}. Все равны. Created once: {checkpoint['creation_timestamp']}. Metadata: DEV\\DEV_UZK_CHECKPOINT.json. Original source size140984, modified2026-10-05T14:28:37.940393+04:00. No original Office open/save.

## 17. CHANGED FILES

Созданы только DEV\\УЗК_DEV.xlsm, DEV\\DEV_UZK_CHECKPOINT.json и audit/evidence файлы в validation\\stage5_uzk_audit_20261007_resume. Там scripts, JSON/CSV, статически извлечённый VBA text, шесть окончательных native PNG pages и report/manifest. Исходники, ВОМ, existingDOCX, DEVВЗМК, profiles/sidecars, Stage1–4 и предыдущий audit не изменены. Нового УЗК.docx нет.

## 18. NOT_RUN

Real registry mark→profile lookup/unsupported-family counts и geometry current4/registry25: недоступен рабочий catalog. Остальные336 DOCX не перечитывались по прямому указанию; классификация остальных first-passport candidates reused prior snapshot. Project-vs-global comparison: exact fallback отсутствует. Native add/remove rows и new document generation/E2E, automatic registry decrement/delete/close, VBA execution, network, productionВЗМК access, АКЗ, Документ о качестве, PDF, EXE, deploy — NOT_RUN. XLS→XLSX conversion, creation/import mark catalog и Stage1–4 changes не выполнялись.

## 19. READY TO IMPLEMENT STAGE 5 — NO

Audit доступных источников завершён. Registry mapping, source date/number и Word row model доказаны; количественный current result=4 marks/33 items. Рабочего project mark catalog нет, поэтому реальные профили/approved-rule compatibility и geometry не установлены. Implementation не начат. После отчёта остановиться.
'''
(OUT/'AUDIT_REPORT.md').write_text(report,encoding='utf-8')
print(json.dumps({'protected':len(before),'classification_counts':dict(scope_counts),'current_marks':4,'Q':33,'source_unchanged':True,'ready':False},ensure_ascii=True))
