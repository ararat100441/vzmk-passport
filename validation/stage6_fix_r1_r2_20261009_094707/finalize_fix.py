"""Final hash binding; manifest is the last written file."""
from pathlib import Path
import json,hashlib,ast,datetime
E=Path(__file__).resolve().parent;R=E.parents[1]
def load(n):return json.loads((E/n).read_text('utf-8'))
def sha(p):return hashlib.sha256(Path(p).read_bytes()).hexdigest()
def save(n,x):(E/n).write_text(json.dumps(x,ensure_ascii=False,indent=2,default=str),encoding='utf-8')
baseline=load('baseline.json');tested=load('tested_product_hashes.json');pair=load('new_pair.json');tests=load('targeted_tests.json')
assert len(tests)==16 and all(t['status']=='PASS' for t in tests)
assert load('audit.json')=={'production':[],'network':[]}
assert load('native_GUI_lock_result.json')['output_empty']
assert not list(E.rglob('.pending-*'))
for kind in ('act','application'):
 native=load('native_'+kind+'.json');after=load('native_'+kind+'.after.json')
 assert native['read_only'] and native['normal_corrupt_load']==0 and native['sheets'][0]['pages']==1
 assert after['unchanged'] and after['before']==after['after']==pair['files'][native['path']]
 assert 'из 1' in (E/('native_'+kind+'_print.txt')).read_text('utf-8')
 assert (E/('native_'+kind+'_print.png')).stat().st_size>10000
current={p:sha(p) for p in baseline['guarded']}
changes={p:{'BEFORE':h,'CURRENT':current[p]} for p,h in baseline['guarded'].items() if current[p]!=h}
assert set(changes)=={str(R/'akz_stage6/core.py'),str(R/'akz_stage6/gui.py')}
assert {p:sha(p) for p in tested['AFTER']}==tested['AFTER']==tested['CURRENT']
assert all(sha(p)==h for p,h in pair['files'].items())
funcs=lambda tree:{n.name:ast.dump(n,include_attributes=False) for n in tree.body if isinstance(n,(ast.FunctionDef,ast.ClassDef))}
x=funcs(ast.parse((E/'BEFORE_core.py').read_text('utf-8')));y=funcs(ast.parse((R/'akz_stage6/core.py').read_text('utf-8')))
unchanged=['number','text','ClimateDay','register_climate','templates','passports','current_passport','catalog','visible_vom','matching','Preview','preview','verify_preview','generate','native_excel']
assert all(x[n]==y[n] for n in unchanged)
save('semantic_scope.json',{'unchanged_AST':unchanged,'calculation_algorithms_changed':False,'pair_publication_architecture_changed':False})
save('native_visual_assessment.json',{'reviewer':'Codex native UI inspection','act':'PASS','application':'PASS','normal_open_no_repair_observed':True,'read_only':True,'close_without_save':True,'pages':{'act':1,'application':1},'application_visible_body_rows':[4,5,6],'application_hidden_body_rows':list(range(7,21)),'total_row_visible':21,'blank_gap':False,'clipping_observed':False,'broken_borders_observed':False,'native_GUI_actionable_Russian_lock_message':'PASS','physical_print':'NOT_RUN'})
save('safety_current.json',{'guarded_count':len(current),'unchanged_count':len(current)-len(changes),'intentional_product_changes':changes,'CURRENT':current})
old={p:h for p,h in current.items() if 'stage6_akz_part1_20261009_085822_release3_TEST_9_2026' in p}
save('old_reviewed_pair_current.json',old)
product={p:sha(p) for p in tested['CURRENT']}
save('tested_product_hashes.json',{'BEFORE':baseline['product'],'AFTER':tested['AFTER'],'CURRENT':product,'AFTER_EQUALS_CURRENT':True})
paths='\n'.join(f'{p}\nSHA-256 {h}' for p,h in pair['files'].items() if p.endswith('.xlsx'))
protected='\n'.join(f'{p}: {h}' for p,h in current.items() if Path(p).name in ('ВЗМК_DEV.xlsx','УЗК_DEV.xlsm','Профили.xlsx','Справочник марок.xlsx','Акт № 4-2026.xlsx','1_2026 - Приложение к АКЗ.xlsx','ВОМ отгрузочный счет 183  (5).xls') and 'sandbox' not in p)
report=f'''1. ACTUAL STATE — FIX COMPLETE; TARGETED VALIDATION PASS
Исходный independent verdict FAIL: только S6-P1-R1 и S6-P1-R2 (MINOR).
До правки 383 fingerprints совпали с exact review baseline и защищёнными старыми DEV outputs.
Git отсутствует. Изменены только akz_stage6/core.py и akz_stage6/gui.py.
Baseline, BEFORE copies и actual exception reproduction сохранены в этой новой папке.

2. ROOT CAUSE R1
write_application очищал A4:G20, но оставлял неиспользуемые Rows.Hidden=False.
На reviewed application строки7:20 действительно пусты и видимы (14 строк).

3. FIX R1
Только output copy: Rows(1:21).Hidden=False; при N<17 Rows(4+N:20).Hidden=True.
Строки не удаляются. Заголовки, N рабочих строк и итог row21 видимы.
validate_saved_pair проверяет сохранённую видимость rows1:21.
C21/E21/G21 остаются SUM(C4:C20)/SUM(E4:E20)/SUM(G4:G20).

4. ROOT CAUSE R2
Эксклюзивный CreateFile share=0 на исправной копии исходного приложения вызывает ambiguous COM error:
Excel сообщает неверный формат/расширение (-2147352567, inner -2146827284).
Ранее GUI показывал str(exc), без Windows-проверки фактического доступа.
Старое поведение повторено до изменения продукта: root_cause_reproduction.json.

5. FIX R2
probe_file_access использует Windows CreateFile с требуемым READ/WRITE доступом.
Windows32/33 => FILE_LOCKED; Windows5 => ACCESS_DENIED; другие ошибки => OTHER_ERROR.
Проверка перед native open и повторно после Excel open/save failure для диагностики гонки.
INVALID_XLSX устанавливается по повреждённому ZIP/CRC, отсутствующим обязательным XML или ошибке XML;
одного COM-текста о формате недостаточно. Исходное исключение сохранено через cause.
GUI показывает русский текст, операцию и полный путь. OUTPUT_ALREADY_EXISTS_STOP и
INPUT_CHANGED_AFTER_PREVIEW_STOP сохранены как эквиваленты OUTPUT_EXISTS / INPUT_CHANGED.
Предварительная проверка не удерживает новый файловый lock во время работы Excel.

6. CHANGED FILES
{chr(10).join(f'{p}: BEFORE {v["BEFORE"]}; CURRENT {v["CURRENT"]}' for p,v in changes.items())}
Полный diff: product_changes.diff. Новые tests/drivers/evidence только в этой validation папке;
новый указатель validation/STAGE6_FIX_CURRENT.txt. Старые evidence/scripts не редактировались.

7. TARGETED VALIDATION R1 — PASS
N1: row4 видима,5:20 скрыты,row21 видима.
N3:4:6 видимы,7:20 скрыты,row21 видима.
N17:4:20 видимы,row21 видима. Все строки и исходный порядок сохранены.
N18:TEMPLATE_CAPACITY_STOP на preview, без усечения или output.
Пары проверены новым validate_saved_pair и независимыми native row/formula/totals reads.
Объединения, границы A1:G21, ширины колонок и print area совпадают с исходным приложением.
Native Print Preview N3: три марки сразу переходят к Итого; одна страница.

8. TARGETED VALIDATION R2 — PASS
Реальный OS handle CreateFile GENERIC_READ,share=0 на healthy second pending XLSX:
FILE_LOCKED с cause winerror32, actionable Russian GUI message; STOP, публикации нет,pending удалён.
После release lock та же preview сформировала полную пару; существующие outputs сохранены.
Повреждённый second pending XLSX: INVALID_XLSX,STOP,полный rollback,не FILE_LOCKED.
Каждый из пяти isolated inputs после preview: исходный INPUT_CHANGED...STOP.
Existing output:OUTPUT_ALREADY_EXISTS_STOP; bytes сохранены.
ACCESS_DENIED: адресный unit с Windows error5; OTHER_ERROR: валидный контейнер/generic failure.
Настоящий Tk messagebox дополнительно открыт и визуально прочитан; native_GUI_lock_*.png/txt/json.
16/16 новых адресных сценариев PASS; прежний unit suite16/16 PASS.
Первый harness запуск остановился на сравнении openpyxl StyleProxy border объектов; исправлено сравнение
копий Border по значениям. Product не менялся. Первый запуск сохранён, suite повторён в sandbox_run2.

9. THREE-MARK REGRESSION — PASS
Источник: original selected VOM,104 visible/397 hidden; matching № АКЗ=numeric0.
Бп2-5:Q1;mass37.20kg;area1.412m2;VOM row40.
Бп2-18:Q1;mass31.32kg;area1.184m2;VOM row53.
Бп2-19:Q1;mass36.74kg;area1.393m2;VOM row54.
TOTAL:3;105.26kg;3.989m2. Calculation algorithms unchanged (semantic_scope.json).
9/2026;05.10.2026;Pirocor Protect057;RAL7004;1layer;80µm;TEST-only preserved.
Real-input preview seal unchanged:e9fc12e4433feafd2982391d9d680838242807d341f425eb653e1afb8e4abbf6.
Официальная выдача STOP; фактические измерения MISSING/TEST placeholders.

10. NEW TEST PAIR
{paths}
PAIR.json SHA-256 {next(h for p,h in pair['files'].items() if p.endswith('PAIR.json'))}
Пара создана новым кодом из original approved inputs и опубликована целиком в новой папке.

11. NATIVE EXCEL VALIDATION — PASS
Microsoft Excel16.0,обе exact новые книги read-only,CorruptLoad0,UpdateLinks0.
Repair/recovery не наблюдались; по одному ожидаемому листу и одной native print page.
Реквизиты/три марки/итоги/TEST поля читаемы; пустого промежутка приложения нет;
обрезание/сломанные границы не обнаружены. Close without save,SHA BEFORE=AFTER.
Native JSON,screenshots и UI text приложены.

12. SOURCE / OLD ARTIFACT SAFETY — PASS
381/381 guarded files неизменны; два изменения продукта намеренные.
Включены все старые DEV output файлы,Stage1–5,DEV registries,profiles,realcatalog,VOM,originaltemplates,
previous implementation/review evidence и manifests. Старые reviewed SHA:
ACT 1a52736427e1529420723ff85f3bfd92340865ece895173ffaaa0dadb48ba06f
APPLICATION 4cd657f84311512255e6a6e39b492c227f18df3ab8b5ab03eb34400f165d9c6d
Повторно проверены CURRENT; точные пути в old_reviewed_pair_current.json.
Ключевые sources:
{protected}
Production registries:0 access attempts; network:0 Python audit attempts,0 network tools.
Python audit охватывает Python open/socket events, это не системная packet capture.

13. EVIDENCE INTEGRITY — PASS
baseline.json связывает exact independent REVIEW_MANIFEST с BEFORE hashes.
AFTER=CURRENT всех7 product/launcher/docs/test файлов. Canonical calculation/selection/preview и
сама generate (включая atomic rename/rollback) AST unchanged.
Manifest записан последним после report/snapshots/evidence; проверен чтением и SHA всех файлов.
Старый PASS не используется как доказательство нового кода: narrow suite/newpair/native UI выполнены заново.

14. REGRESSION IMPACT
Только видимость output-строк и диагностика файловых операций/GUI.
Расчёты,источники,нумерация,test-only guard,capacity17 и atomic rename публикации не менялись.
Узкие тесты/прежний Stage6 unit suite PASS; полный regression Stage1–5 не выполнялся.

15. NOT_RUN
Stage7;АКЗ1/2/3;production access/output;официальная выдача;реальные measurements/climate engine;
физическая печать;полный Stage1–5 regression;real ACL mutation для ACCESS_DENIED;системный network capture.
Расширение шаблона свыше17 строк не выполнялось. Codebase index tools отсутствуют.

16. READY FOR TARGETED WORK REREVIEW — YES
S6-P1-R1/S6-P1-R2 исправлены и адресно проверены. Независимый rereview ещё не выполнен.
Работа остановлена в границах consolidated fix.
'''
(E/'FIX_REPORT.txt').write_text(report,encoding='utf-8')
files={str(p):sha(p) for p in E.rglob('*') if p.is_file() and p.name!='EVIDENCE_MANIFEST.json'}
manifest={'created_at':datetime.datetime.now().astimezone().isoformat(),'scope':['S6-P1-R1','S6-P1-R2'],'ready_for_targeted_rereview':True,'review_baseline':baseline['review_manifest'],'review_baseline_sha256':sha(baseline['review_manifest']),'product':load('tested_product_hashes.json'),'guarded_current':current,'new_pair':pair,'tests':{'targeted':16,'unit':16},'files':files,'written_last':True}
save('EVIDENCE_MANIFEST.json',manifest)
assert all(sha(p)==h for p,h in load('EVIDENCE_MANIFEST.json')['files'].items())
assert all(sha(p)==h for p,h in current.items())
print('FINAL PASS',len(files),'evidence files; manifest SHA256',sha(E/'EVIDENCE_MANIFEST.json'))
print('REPORT',str(E/'FIX_REPORT.txt'))
