"""Finish the narrow User E2E correction; final evidence manifest written last."""
from pathlib import Path
from datetime import datetime,timezone
import difflib
import sys
sys.dont_write_bytecode=True
HERE=Path(__file__).resolve().parent;ROOT=HERE.parents[1];sys.path.insert(0,str(ROOT))
from stage7_quality.provenance import read,write,verify,sha,now
from stage7_quality.core import verify_manifest,code_bindings
from stage7_quality.word import parse,NS,canonical,table_cells,text
from stage7_quality.presentation_revision import VERSIONS


def link(path,label=None):
    p=Path(path).resolve();return '['+(label or p.name)+'](<'+str(p).replace('\\','/')+'>)'


pre=read(HERE/'PREFLIGHT.json');selected=pre['manifest'];m=read(selected['path'])
verify_manifest(m);verify({selected['path']:selected['sha256']})
protected=read(HERE/'PROTECTED_BEFORE.json');verify(protected)
product_before=read(HERE/'PRODUCT_BEFORE_FINAL_TESTS.json');verify(product_before)
assert product_before==read(HERE/'PRODUCT_AFTER_FINAL_TESTS.json')==code_bindings()
result=read(HERE/'NEW_QUALITY_RESULT.json');receipt=read(result['receipt']['path'])
assert receipt['output_sha256']==sha(result['output'])==result['output_sha256']
assert read(HERE/'TARGETED_TEST_RESULTS.json')['status']=='PASS'
assert read(HERE/'GUI_SMOKE.json')['status']=='PASS'
assert read(HERE/'POINT12_RANGE_VISUAL_COMPARISON.json')['pixel_equal'] is True
first=ROOT.parent/'Гараж казахстан/Паспорт/1_2026 - Документ о качестве.docx'
manual=Path(pre['manual_reference']['path'])
# CPython read-only file handles allow the reference already open in Word.
# No attach/close/save of the user's Word session is performed.
assert sha(manual)==pre['manual_reference']['sha256']
assert sha(first)==pre['template_sha256']
write(HERE/'PROTECTED_AFTER.json',{p:sha(p) for p in protected})
write(HERE/'SOURCE_INTEGRITY.json',{'status':'PASS','protected_files':len(protected),'changed':0,
    'stage1_6_changed':0,'old_outputs_changed':0,'old_receipts_changed':0,'original_manifest_changed':0,
    'manual_reference_sha256':sha(manual),'first_template_sha256':sha(first),
    'current_product_bindings':len(product_before),'before_after_current_equal':True,
    'generator_reruns':{'VIK':0,'UZK':0,'AKZ':0},'production_access':'NOT_RUN','pdf':'NOT_RUN','exe':'NOT_RUN','deployment':'NOT_RUN','stage8':'NOT_RUN'})
diff=[]
changed=[]
for name in VERSIONS:
    before=HERE/'code_before'/name;after=ROOT/'stage7_quality'/name
    diff.extend(difflib.unified_diff(before.read_text(encoding='utf8').splitlines(True),after.read_text(encoding='utf8').splitlines(True),fromfile='accepted_before/stage7_quality/'+name,tofile='current/stage7_quality/'+name))
    changed.append({'path':str(after),'kind':'MODIFIED','before':sha(before),'after':sha(after)})
(HERE/'EXACT_PRODUCT_DIFF.diff').write_text(''.join(diff),encoding='utf8')
added=ROOT/'stage7_quality/presentation_revision.py'
changed.append({'path':str(added),'kind':'ADDED','sha256':sha(added)})
write(HERE/'CHANGED_FILES.json',{'product':changed,'other_existing_product_changes':[],
    'targeted_tests':str(HERE/'targeted_tests.py'),'new_evidence_directory':str(HERE),
    'protected_sources_changed':[],'old_receipts_and_outputs_changed':[]})
for label,path in [('FIRST',first),('NEW',Path(result['output']))]:
    root=parse(path);rows=root.findall('w:body/w:tbl',NS)[0].findall('w:tr',NS)[40:49]
    with (HERE/('POINT12_'+label+'_XML.xml')).open('xb') as f:f.write(b'<protected>'+b''.join(canonical(r) for r in rows)+b'</protected>')
assert sha(HERE/'POINT12_FIRST_XML.xml')==sha(HERE/'POINT12_NEW_XML.xml')
cells=table_cells(parse(result['output']));point13=text(cells[0][50][0]);point14=[text(p) for p in cells[0][53][0].findall('w:p',NS)]
native=read(read(receipt['validation']['path'])['checks']['native']['evidence'])
report=[
'# ВЗМК Этап 7 — точечное исправление после User E2E',
'',
'READY FOR TARGETED WORK 07 REVIEW: YES. BLOCKER0. Только DEV/TEST, official issuance запрещена. После отчёта STOP.',
'',
'## 1. ROOT CAUSE',
'',
'Представление пунктов13/14 использовало status-bearing display_title из manifest; пункт14 также выводил приложение АКЗ отдельной записью и служебное ограничение SCOPE. Шаблонное значение договора сохраняло собственное № рядом с постоянной меткой «Договор №». User E2E утвердил окончательное представление по ручному9/2026. Исправлена только эта логика; исходный Stage7 audit и этапы1–6 не повторялись.',
'',
'## 2. APPROVED USER RULES',
'',
'Первый паспорт: SHA `'+pre['template_sha256']+'`; ручной9/2026: SHA `'+pre['manual_reference']['sha256']+'`. Оба проверены ДО чтения. Пункт12 целиком из первого, включая «от окалины ржавчины по ГОСТ9.402» без добавления союза. Пункты13/14 — один последовательный абзац из четырёх актов через запятую; в14 после него точная фраза хранения. AKZ-app проверяется в manifest как часть пары, в Word не перечисляется. TEST header остаётся, служебные статусы/неполнота контроля остаются в metadata/отчёте/GUI. Экземпляры не выдуманы.',
'',
'## 3. EXACT CODE CHANGES',
'',
'Изменены только stage7_quality/word.py, core.py, provenance.py; добавлен stage7_quality/presentation_revision.py. '+link(HERE/'EXACT_PRODUCT_DIFF.diff')+' и '+link(HERE/'CHANGED_FILES.json')+'.',
'',
'word.py: references формирует чистые названия актов по type + проверенным number/date, отделяя машинный status от текста Word. Число/год/даты будущего паспорта не зашиты в presenter; in-memory27/2030/18.07.2030 проверен. acts_paragraph использует запятые. AKZ_APPENDIX отфильтрован только на этапе представления; проверка обязательной пары сохранена. STORAGE приведён к точному тексту пользователя. SCOPE не вставляется в13/14. contract_value снимает только лишний начальный № из значения, постоянная метка не изменяется. validate_quality проверяет новые точные абзацы и договор.',
'',
'core.py/provenance.py: необходимая привязка новой версии генератора при повторном использовании прежнего frozen manifest. Старые receipts сохраняют SHA своего исторического Stage7 wrapper. Новый helper разрешает только ТРИ точных пары old→new SHA в этом исправлении и требует побайтные BEFORE-копии, сохранённые до правок. Остальные code hashes, все inputs/output SHA и START/RESULT/VALIDATION цепочки проверяются строго. Изменённый архив либо неизвестный old/new SHA дают STOP. Новый quality receipt фиксирует CURRENT code bindings и архивные proof inputs, не подменяет старые receipts. '+link(HERE/'AUTHORIZED_CODE_TRANSITION.json')+'.',
'',
'## 4. POINT 12 INTEGRITY',
'',
'PASS. Канонический XML всех строк41–49 полностью равен первому паспорту, включая текст, пробелы, орфографию, pPr/rPr, маркированные списки, ячейки, границы и объединения. Native Word Range первого[2036,2702), нового[2077,2743) имеет одинаковый текст и координаты:620.6500–733.6500pt, страница1. Pirocor Protect057,не менее80мкм,пескоструйная обработка,степень2 сохранены. Прямые Range.EnhMetaFileBits→PNG блоков равны попиксельно: '+link(HERE/'POINT12_RANGE_VISUAL_COMPARISON.json')+'.',
'',
'Первичное сравнение кропов полной страницы выявило177отличающихся pixels из1341680 в одном контуре glyph; эти диагностические файлы сохранены. Полные страницы визуально просмотрены, координаты/структура равны, прямой native render защищённого диапазона подтвердил pixel_equal=true. Никакой правки текста/шрифта пункта12 для сравнения не выполнялось.',
'',
'## 5. POINT 13 RESULT',
'',
point13,
'',
'Один абзац, точное совпадение с ручным9/2026 по тексту и semantic pPr/rPr. [ТЕСТ]/SAMPLE_ONLY/АКЗ№0 в названия не вставлены. Номера/даты получены из unchanged verified selected records.',
'',
'## 6. POINT 14 RESULT',
'',
point14[0],
'',
point14[1],
'',
'Два абзаца, точное совпадение с ручным9/2026 по тексту и semantic pPr/rPr. Сохранено исходное пояснение «(перечень документов с указанием числа экземпляров)». Отдельной строки AKZ-app и служебных пояснений нет. Приложение1 остаётся в пункте2. Обязательные файлы пары АКЗ остаются в manifest и проходят проверку SHA/parent/pair_id.',
'',
'## 7. CONTRACT NUMBER CHECK',
'',
'Постоянная метка «Договор №» сохранена. Значение «   06/26 от03.06.2026» соответствует ручному образцу; фактический номер и дата договора сохранены. Двойного № в результате нет. Targeted test сравнил label/value с ручным XML и проверил динамическое значение другого договора.',
'',
'## 8. TEST STATUS PRESERVATION',
'',
'В заголовке видны «ТЕСТ / ЧЕРНОВИК. НЕ ДЛЯ ОФИЦИАЛЬНОЙ ВЫДАЧИ.». Unchanged manifest сохраняет TEST/SAMPLE_ONLY, official_confirmed=false, AKZ0 assigned3marks/3pieces. Новый receipt:results_confirmed=false,official_issue=false,statusTEST. ДНО/годен остаются образцом УЗК; реальные замеры АКЗ не подтверждены,3шт.не подтверждают всю отгрузку288шт. Подписи/печать не созданы, пустые места и М.П. сохранены. GUI callback smoke подтвердил прежние предупреждения и ограниченный охват; generator_calls0 без подтверждения. '+link(HERE/'GUI_SMOKE.json')+'.',
'',
'## 9. TARGETED TESTS',
'',
'27/27 PASS; failures0/errors0/skipped0. '+link(HERE/'TARGETED_TEST_RESULTS.json')+'; '+link(HERE/'TARGETED_TESTS.log')+'. Позитивные проверки:point12 exact XML/format;13/14 manual text/paragraph/run properties;future number/date presenter;noAKZ-app/noadminlines;contract;TEST;native2A4pages;current code/output SHA/manifest binding;opaque parts/signature blanks. Отрицательные:missing act,missing AKZ-app,wrong status,wrong date,changed SHA в isolated copy,unknown code SHA,changed archived code snapshot,protected-block formatting mutation,output collision. Только isolated fixtures/in-memory; fullStage1–6 regressionNOT_RUN.',
'',
'## 10. NATIVE WORD VALIDATION',
'',
'Новый DOCX:normal read-only Microsoft Word COM,OpenAndRepair=False,Close(False),before/afterSHAequal. Две страницы A4:595.30×841.90pt;1section,2tables,14logicalpoints. Numbering/MERGEFIELD/opaque ZIP parts/headers/footers/signature blanks сохранены. Видимая TEST-плашка сохранена;13/14 без служебных вставок, договора№№нет, третьейстраницы/обрезки/перекрытийнет. Native PNG обеих страниц и первого/ручного references сохранены. '+link(HERE/'VISUAL_REVIEW.json')+'. PDF не экспортировался. Ручной источник был открыт пользователем в Word:прочитан только readonly shared handle, native comparison выполнено по побайтному SHA-verified snapshot; пользовательский Word не закрывался и не сохранялся.',
'',
'## 11. NEW DOCX PATH / SHA',
'',
link(result['output'])+' — SHA-256 `'+result['output_sha256']+'`.',
'',
'Новая уникальная папка QUALITY_TEST_V2. Предыдущий Stage7 DOCX не перезаписан и не выдан за результат исправленного генератора.',
'',
'## 12. RECEIPT / MANIFEST INTEGRITY',
'',
'Новый '+link(result['receipt']['path'])+' — SHA `'+result['receipt']['sha256']+'`; operation_id `'+receipt['operation_id']+'`; assembly_run_id `'+receipt['assembly_run_id']+'`. START→RESULT→VALIDATION реального вызова только stage7_quality.word.write_quality. Новый receipt COMPLETE_VALIDATED_TEST и current generator hashes; сохранены архивные SHA исходного producer wrapper.',
'',
'Использован существующий '+link(selected['path'])+' — SHA `'+selected['sha256']+'`, без переписывания или reseal. Все6записей, собственные source run IDs/receipts, номера, даты, пути и SHA повторно проверены. ВИК/УЗК/АКЗ не запускались повторно. Изначальная и новая quality версии имеют разные output SHA; старые receipt chains сохранены.',
'',
'## 13. SOURCE INTEGRITY',
'',
f'PASS: {len(protected)} защищённых файлов BEFORE=AFTER, changed0; все source registries/VOM/catalog/profiles/original DOCX/old outputs/receipts/manifest/evidence сохранены. Current bound product {len(product_before)} файлов BEFOREfinaltests=AFTER=CURRENT. Разрешённые3Stage7changes отделены от protected source snapshot. '+link(HERE/'SOURCE_INTEGRITY.json')+'. Production/web/registry writes/UZK decrement/PDF/EXE/deploy/Stage8NOT_RUN.',
'',
'## 14. CHANGED FILES',
'',
'3modified Stage7files +1new presentation_revision helper; новые targeted tests и evidence только в '+str(HERE)+'. Изменений Stage1–6,DEV registry,production registry,UZK registry,VOM,catalog,profiles,originalWord,existingTESToutputs,previousreceipts нет. '+link(HERE/'CHANGED_FILES.json')+'. Индексные инструменты index_status/codebase_context вернули INTERNAL_ERROR; выполнены только точечные чтения известных файлов.',
'',
'## 15. REMAINING FINDINGS',
'',
'BLOCKER0 для targeted Work07 review. Официальная выдача по-прежнему запрещена:фактические результаты УЗК/АКЗ не подтверждены. Новая ручная User E2E приёмка и независимое targeted Work07 review NOT_RUN. Full regression1–6 и повторные generatorsNOT_RUN по scope. Полные page PNG проверены визуально, диагностическое небольшое отличие полногоpageGDIкропа не затронуло protected XML/text/positions; прямойnativeRangepixelcomparisonPASS.',
'',
'## 16. READY FOR TARGETED WORK 07 REVIEW: YES',
'',
'Точечное исправление завершено. EVIDENCE_MANIFEST.json записан последним, затем проверен read-only. После отчёта STOP; Stage8/EXE/deployment не запускались.',
]
(HERE/'TARGETED_USER_E2E_REPORT.md').write_text('\n'.join(report)+'\n',encoding='utf8')
verify(protected);verify(product_before);verify_manifest(m)
evidence={str(p):{'sha256':sha(p),'bytes':p.stat().st_size} for p in sorted(HERE.rglob('*')) if p.is_file() and p.name!='EVIDENCE_MANIFEST.json'}
final={'schema':'STAGE7_TARGETED_USER_E2E_EVIDENCE_V1','created_at':now(),'status':'READY_FOR_TARGETED_WORK07_REVIEW',
    'evidence':evidence,'product_changes':changed,'current_product':product_before,'original_manifest':selected,
    'new_receipt':result['receipt'],'new_output':{'path':result['output'],'sha256':result['output_sha256']},
    'targeted_tests':27,'protected_sources':len(protected),'protected_changed':0,'official_issue':False,'stage8':'NOT_RUN'}
write(HERE/'EVIDENCE_MANIFEST.json',final)
verify({p:v['sha256'] for p,v in evidence.items()});verify(product_before)
print({'status':'PASS','report':str(HERE/'TARGETED_USER_E2E_REPORT.md'),'evidence_files':len(evidence),
       'manifest_sha256':sha(HERE/'EVIDENCE_MANIFEST.json'),'protected_files':len(protected),
       'product_bound_files':len(product_before),'ready_for_targeted_work07_review':True})
