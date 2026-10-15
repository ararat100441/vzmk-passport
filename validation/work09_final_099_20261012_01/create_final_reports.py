"""Produce final review documents; no product or runtime actions."""
import hashlib,json,os
from pathlib import Path
from datetime import datetime,timezone
OUT=Path(__file__).resolve().parent;ROOT=OUT.parents[1]
def read(p):return json.loads(Path(p).read_text(encoding='utf-8-sig'))
def sha(p):return hashlib.sha256(Path(p).read_bytes()).hexdigest()
def save(name,obj):
    with (OUT/name).open('x',encoding='utf8') as f:
        json.dump(obj,f,ensure_ascii=False,indent=2);f.flush();os.fsync(f.fileno())
def doc(name,text):
    with (OUT/name).open('x',encoding='utf8') as f:f.write(text);f.flush();os.fsync(f.fileno())
def link(label,path):return '['+label+'](<'+str(path).replace('\\','/')+'>)'
def evidence(name):return link(name,OUT/name)
exe=ROOT/'stage9_dist_099/VZMK-Passport-0.9.9/VZMK-Passport-0.9.9.exe'
expected='21ff2f7cc0d4fca746094c46d3594ae0b1dbf991c40266195e22c1030c0ccadc'
assert sha(exe)==expected
p=read(OUT/'FINAL_PROCESS_AUDIT.json'); i=read(OUT/'FINAL_IMMUTABLE_INTEGRITY.json')
assert p['office_count']==57 and p['job_count']==26 and i['status']=='PASS'
not_run=[
 {'check':'Visible frozen GUI first open / repeat open / passport 9-10-9 switching','status':'NOT_RUN','reason':'SAFETY_STOP after unexpected non-review Excel exit'},
 {'check':'Visible GUI Word.Close failure, no false SUCCESS, explicit subsequent Open','status':'NOT_RUN','reason':'SAFETY_STOP; source mocks are separate'},
 {'check':'First GUI shutdown with viewer; fresh GUI launch and shutdown','status':'NOT_RUN','reason':'SAFETY_STOP'},
 {'check':'GUI TEST generation, counters, closure and recovery','status':'NOT_RUN','reason':'SAFETY_STOP; 26 actual frozen CLI tasks passed their expectations'},
 {'check':'Additional frozen lifecycle normal/error/cleanup-error/crash/restart with Office alive','status':'NOT_RUN','reason':'SAFETY_STOP; transaction during_publish worker exit73 was executed, but is a different case'},
 {'check':'Additional frozen diagnostic collision and four repeated native verifier protocol runs','status':'NOT_RUN','reason':'SAFETY_STOP; three native verifier tasks already executed within E2E'},
 {'check':'Full 17-case frozen negative suite; normal overlap, overlap with crash, controlled foreign sentinels and forged ownership','status':'NOT_RUN','reason':'SAFETY_STOP; source ownership/window guards and DEV9 closure denial executed separately'},
 {'check':'Continuous non-review Office session preservation / cause of non-review exit / historical Word3132 cause','status':'NOT_VERIFIED / UNKNOWN','reason':'No attribution evidence; no unknown process was closed by review'},
 {'check':'Visual inspection of every generated page and all AKZ1/2/3 variants; physical network disconnection','status':'NOT_RUN','reason':'Bounded review; AKZ0 and two native quality page images inspected'},
 {'check':'Separate per-document Stage7 receipt chain for the Stage9-generated bundle','status':'NOT_PRODUCED / NOT_CLAIMED','reason':'This path produces its own sealed Stage9 TEST manifest and Stage8 transaction records'},
 {'check':'Production operation, real UZK deduction, official closure, deployment','status':'NOT_RUN','reason':'Explicitly prohibited by user'}]
save('NOT_RUN.json',{'runtime_stop':str(OUT/'FOREIGN_OFFICE_STOP.json'),'checks':not_run})
save('REVIEW_METHOD.json',{'review_type':'independent targeted review','fresh_runtime':'exact frozen EXE0.9.9,26 CLI tasks; no Python in EXE PATH','source_tests':'17 existing targeted tests + 15 independent window/ownership guard cases +3 mocked viewer cycles','source_to_EXE_binding':'49 compiled modules;60 raw files, only co_filename normalized','historical_self_evidence':'SHA readback only; not fresh independent execution','codebase_index':'index_status and codebase_context returned INTERNAL_ERROR earlier in this review; bounded addressed source reads used as fallback','GUI_authorized_by_latest_user_request':True,'GUI_not_run_reason':'new safety stop, not missing GUI authorization','raw_inherited_label':'e2e RESULT.gui_smoke=NOT_RUN_USER_STOP retained as originally written by copied protocol; current status is NOT_RUN_SAFETY_STOP','after_stop_actions':'read-only evidence, own evidence writes, query/synchronize process handles; no new Office or GUI launches','no_product_fixes':True,'no_subagents':True})
old=read(OUT/'CONTROLLER_PROVENANCE.json')
save('CONTROLLER_SHA256_SEALED.json',{'files':{str(x):sha(x) for x in OUT.glob('*.py')},'copy_provenance':[x|{'current_copy_sha256':sha(x['copy']),'changed_after_initial_copy':sha(x['copy'])!=x['copy_sha256']} for x in old['adapted']],'prepared_NOT_RUN':read(OUT/'CONTROLLER_SHA256_FINAL.json')['prepared_but_unexecuted'],'note':'This map includes final audit/report/seal controllers; earlier FINAL map is retained as its earlier snapshot.'})
save('CODEX09_FINDINGS.json',{'BLOCKER':0,'MAJOR':0,'MINOR':0,'new_reproduced_product_findings':[],'previous_major':'source and compiled binding verified; independent frozen GUI acceptance NOT_RUN','previous_minor':'source and compiled binding verified; independent frozen Word.Close fault acceptance NOT_RUN','observation':'Unexpected non-review Excel18392 and39156 exits after all own jobs; cause UNKNOWN. No source defect or responsible actor established. Stop is a coverage/safety boundary, not an invented product finding.'})
result={'review':'WORK09_FINAL_099','completed_artifact_set':True,'runtime_execution_completed':False,'exe':str(exe),'exe_sha256':expected,'source_integrity':'PASS','protected_files':67804,'source_tests':17,'ownership_window_source_cases':15,'mock_viewer_cycles':3,'frozen_CLI_tasks':26,'frozen_CLI_expectations':'24 PASS + 2 expected FAIL;26/26 expectations met','own_Office_instances':57,'own_orphans':0,'normal_force':0,'all_force':0,'validated_WM_CLOSE':2,'physical_documents':18,'logical_records':21,'GUI_E2E':'NOT_RUN_SAFETY_STOP','MAJOR_fix':'VERIFIED_SOURCE_AND_COMPILED_BINDING; FROZEN_GUI_NOT_RUN','MINOR_fix':'VERIFIED_SOURCE_AND_COMPILED_BINDING; FROZEN_WORD_CLOSE_FAULT_NOT_RUN','BLOCKER':0,'MAJOR':0,'MINOR':0,'FINAL_VERDICT':'FAIL','verdict_reason':'Required independent GUI and remaining frozen fault protocols are NOT_RUN after mandated safety STOP. No newly reproduced product defect is claimed.','READY_FOR_USER_E2E':'NO','READY_FOR_PRODUCTION':'NO','stop_evidence':str(OUT/'FOREIGN_OFFICE_STOP.json'),'evidence_manifest':str(OUT/'EVIDENCE_MANIFEST.json'),'report':str(OUT/'WORK09_REPORT.md'),'final_actions':'STOP; no product change, deployment, official closure or real deduction'}
save('REVIEW_RESULT.json',result)
doc('CODEX09_FINDINGS.md',f'''# Work09 final independent review — EXE0.9.9

**Подтверждённые новые дефекты: BLOCKER0 / MAJOR0 / MINOR0. FINAL VERDICT: FAIL из-за незавершённого обязательного покрытия.** Отсутствие новых findings не означает полный PASS.

W09-097-M01: прежняя двойная эксклюзивная запись устранена в текущем source: ownership.py:293–311 — unconditional handle.Close, затем одна final save:304; diagnostics.py:12–30 — уникальные viewer/event пути с ограниченным retry, без overwrite. 17 source-тестов и три mock-цикла прошли; compiled/raw binding к точной EXE подтверждён. Повторное открытие, переключение и первое закрытие настоящего frozen GUI **NOT_RUN**, поэтому независимая приёмка MAJOR остаётся открытой.

W09-097-m01: office.py:42–64 независимо выполняет Word.Close/Word.Quit и Excel.Close/Excel.Quit, сохраняет structured errors; diagnostics.py:8–11 включает HRESULT/excepinfo/traceback; ownership.py:300–310 сохраняет первичную ошибку и secondary errors. gui.py:175–191 показывает FAIL и прекращает новый Open; shutdown:247 учитывает результат. Source-проверки прошли; реальный frozen GUI Word.Close fault **NOT_RUN**. Независимая runtime-приёмка MINOR остаётся открытой.

Новая остановка: Excel PID18392, creation_time2026-10-12T11:56:13.526000+00:00, родитель explorer.exe PID17084, exit0 зафиксирован в11:56:16UTC; ещё один неизвестный Excel39156 вышел0 в11:57:13UTC. Все собственные26 задачи закончились в11:52:19.957UTC; эти identities отсутствуют в собственных57 START/created записях. Причина закрытия **UNKNOWN**, причинная связь с продуктом не доказана. Историческая причина закрытия Word3132 также UNKNOWN. Work не завершал неизвестные Office-процессы.

Это observation и обязательный STOP по подготовленному заданию; точного дефектного файла/строки или минимального продуктового исправления для UNKNOWN нет. Work ничего не исправлял. {evidence('FOREIGN_OFFICE_STOP.json')}, {evidence('FINAL_PROCESS_AUDIT.json')}, {evidence('NOT_RUN.json')}.
''')
doc('USER_E2E.md',f'''# Короткий оставшийся сценарий независимого GUI E2E — НЕ ВЫПОЛНЕН

Сначала отдельно разобрать текущий safety STOP и согласовать безопасное окно теста. Чужой Office не закрывать; UNKNOWN ownership означает STOP. Ни один пункт ниже не отмечен PASS. Использовать exact {link('EXE0.9.9',exe)}, SHA `{expected}`, весь onedir и новый изолированный DEV/TEST workspace. Рабочий ВЗМК.xlsx и принятые outputs не использовать для записи.

1. Записать baseline Office PID/creation time/image. Из sealed DEV-копии создать новый DEV-стенд, выбрать TEST_VOM.xls, проверить9/2026 и явно подтвердить TEST-комплект. Сверить6 физических файлов /7 records, два ВИК в одном DOCX, УЗК SAMPLE_ONLY, выбранный АКЗ0, качество последним.
2. Нажать «Открыть Word/Excel», затем повторить два раза подряд. Каждый цикл должен дать новый diagnostic UUID без FileExistsError, только3 DOCX и3 XLSX ReadOnly. Сверить SHA до/после. Переключить TEST-комплекты9→10→9 и повторить открытие.
3. На доказанном собственном Word текущего viewer через обычное закрытие закрыть его документы. Нажать повторный Open в ВЗМК: Word.Close failure должен сохраниться с HRESULT/traceback и diagnostic path; Excel cleanup должен завершиться отдельно; новый viewer и SUCCESS запрещены. После acknowledgement новый явный Open проверить отдельно. Не закрывать Word по одному PID без birth/image/retained-handle ownership.
4. С активным собственным viewer закрыть ВЗМК один раз: GUI и собственный Office должны завершиться штатно, force0/orphan0. Свежий запуск и первое закрытие повторить.
5. Только на новом artificial8001/2099 fixture проверить GUI backup/TEST-закрытие и повтор: [12,10]→[9,6], затем без второго списания. На отдельном искусственном interrupted fixture проверить восстановление: [12,10], exact seed XLSM SHA, все11 VBA/OLE streams. Проверить STOP для DEV9/2026. Официальная выдача и реальные списания запрещены.
6. Сохранить screenshots, исходные errors/results, manifests/journals, Office identities, EXE/source/product SHA до/после; запечатать новую evidence-серию. Не дописывать runtime в текущий sealed каталог.
''')
doc('WORK09_REPORT.md',f'''# ВЗМК — Этап9 — Final independent review EXE0.9.9

12.10.2026, Europe/Samara. **FINAL VERDICT: FAIL. READY FOR USER E2E: NO. READY FOR PRODUCTION: NO.** Подтверждённых новых дефектов: **BLOCKER0 / MAJOR0 / MINOR0**. Причина FAIL — обязательные независимые GUI/часть frozen fault-проверок NOT_RUN после safety STOP; полного заключения об отсутствии регрессий нет. Это завершённый отчёт о фактически выполненной серии, а не полный runtime PASS. Work не менял product.

## 1. ACTUAL EXE / SHA

{link('VZMK-Passport-0.9.9.exe',exe)}

SHA-256 до/после: **`{expected}`**, совпадает с expected. Выполнялась только0.9.9. Onedir **1117 файлов**, aggregate `0a9abec9d2583aaf4370eb9aa143d9bed63ebd7561b50b4df497ce84f29cad01`. Source aggregate `913ce2bc334ed2dbfae70311089000bf0b47220b5caef3b591e4175b032c983c`. **49 compiled modules /60 raw source copies** соответствуют текущим исходникам; нормализован только co_filename. Запуски EXE использовали PATH System32/Windows без Python. {evidence('BUILD_SHA256_AFTER.json')}, {evidence('COMPILED_SOURCE_AFTER.json')}.

## 2. PREVIOUS MAJOR — W09-097-M01

**Source/compiled verification PASS; независимая frozen GUI приёмка NOT_RUN.** В ownership.py:293–311 handle.Close выполняется безусловно, final ownership.json создаётся один раз после cleanup (304). Exclusive creation сохранён. diagnostics.py:12–30 выделяет UUID viewer/event paths с ограниченным retry без overwrite. Source-тесты проверили повторный teardown, final-save collision, обязательное освобождение handle, primary preservation, уникальность paths. Три source/mock viewer цикла дали3 уникальных diagnostics без FileExistsError. Это не реальные GUI повторы. {evidence('SOURCE_TARGETED_17.json')}, {evidence('VIEWER_REPEAT_SOURCE.json')}.

## 3. PREVIOUS MINOR — W09-097-m01

**Source/compiled verification PASS; реальный frozen Word.Close fault NOT_RUN.** office.py:42–64 независимо выполняет обе cleanup-ветки и сохраняет structured cleanup_errors; diagnostics.py:8–11 содержит step/type/HRESULT/excepinfo/traceback. ownership.py:300–310 сохраняет primary и добавляет secondary errors. gui.py:175–191 учитывает FAIL, показывает diagnostic path, не запускает новый viewer и не сообщает SUCCESS; shutdown:247 учитывает результат. Source-тесты Word.Close/HRESULT, Excel continuation, primary-preservation и GUI callbacks на fake widgets прошли. Настоящий GUI/COM fault не подменён ими.

## 4. INDEPENDENT GUI E2E

**NOT_RUN_SAFETY_STOP.** В этой независимой серии настоящий GUI ВЗМК не запущен; первое/повторное открытие viewer, GUI9→10→9, Word.Close failure, единый интерфейс, счётчики и первое закрытие GUI не подтверждены. Причина — новый runtime safety STOP, а не отсутствие разрешения GUI в последнем запросе пользователя.

Переданная Codex GUI self-evidence проверена чтением и SHA, но не считается независимым GUI E2E:2160 resume references и3606 GUI references, mismatch0. Историческое поле `gui_smoke=NOT_RUN_USER_STOP` в raw CLI RESULT сохранено без переписывания; фактический текущий статус указан здесь и в REVIEW_RESULT/NOT_RUN. {evidence('HANDOFF_READBACK.json')}, {evidence('REVIEW_METHOD.json')}, {evidence('USER_E2E.md')}.

## 5. OFFICE PROCESS SAFETY

В выполненном fresh CLI scope: **26 jobs /57 собственных Office**, **normal force0 /all force0 /own orphan0 /own unknown identities0**. Каждый сопоставлен с created/START и монитором по **PID +creation time +image +launcher parent**, а не только PID; записи retained CREATE_PROCESS handles совпали. Все57 вышли0. **2 validated EMPTY_EXCEL_WM_CLOSE_BEGIN**, только собственный пустой Excel после Quit, без COM refs. **17 source-тестов +15 ownership/window cases PASS** покрыли чужого launcher, not-owned/not-child, PID/birth/image/handle mismatch и оконные гонки. Контролируемые foreign sentinels/forged ownership и lifecycle crash с ещё живым Office **NOT_RUN**.

После окончания всех jobs в **11:52:19.957UTC /15:52:19.957Samara** появился не принадлежащий review Excel18392, birth11:56:13.526UTC; выход0 замечен11:56:16.092UTC. Родитель explorer.exe17084. Excel39156 также вышел0 после задач. Эти identities не совпадают ни с одним собственным START/created. Причины **UNKNOWN**, вина продукта не установлена. Word3132 historical cause остаётся UNKNOWN. Подготовленное задание прямо требует: **«unexpected non-test Office exit → немедленно STOP и диагностика»**; новые runtime-запуски не выполнялись после обнаружения STOP при возобновлении. Монитор только QUERY/SYNCHRONIZE, без termination API; его handles освобождены. Чужие/неизвестные Office не закрывались; global taskkill отсутствовал. {link('Точное условие задания',ROOT/'validation/stage9_gui_099_20261012_01/WORK09_TARGETED_ASSIGNMENT.md')}, {evidence('FOREIGN_OFFICE_STOP.json')}, {evidence('POST_STOP_PROCESS_SNAPSHOT.json')}, {evidence('FINAL_PROCESS_AUDIT.json')}.

## 6. FROZEN EXE INTEGRATION — Stage1–8 boundary

**26/26 ожидаемых исходов PASS в CLI scope:24 PASS +2 ожидаемых FAIL.** Создан новый отдельный workspace `C:\\Users\\arara\\W09F099_7e1d54aa\\163fa85fa0b44845bf8d50e34856fcde`. Полный аудит Stage1–8 не повторялся. Реальная EXE импортировала sealed DEV-источники, проверила/сформировала/native-verified **9/2026 → новый10/2026 →9/2026**. Native verifier использовал настоящие Word/Excel; это CLI, не GUI Viewer.

Каждый комплект: **104 visible /397 hidden excluded /288шт. /15822.14кг /1913.494м² /УЗК33**, Word3 /Excel3. ВИК9 и9-1 находятся в одном DOCX; УЗК SAMPLE_ONLY, выбранный АКЗ0 и приложение; документ качества последний. Итого **18 физических файлов /21 логическая запись**, по6/7 на комплект. Numbers/dates, M1/connections и quality structural invariants проверены. {evidence('EXE_E2E_FINAL.json')}, {evidence('DOCUMENT_AND_XLSM_AUDIT.json')}.

## 7. MANIFEST / RECEIPTS

**PASS для фактически созданных Stage9 manifests и Stage8 transaction records.** Seals, input/document/native evidence SHA и journal hash chains валидны.9 sealed request bindings на artificial fixture. Separate per-document Stage7 receipt files этот Stage9 path не создаёт; независимый PASS такой отсутствующей цепочке не заявлен. {evidence('MANIFEST_AND_TRANSACTION_RECORDS_AUDIT.json')}.

## 8. BACKUP / TEST CLOSE / ONCE-ONLY / RESTORE

**PASS actual frozen CLI.** Искусственный паспорт8001/2099: INTENT→BACKUP_VERIFIED→STAGED→APPLYING→COMMITTED; CLOSED_TEST, остатки **[12,10]→[9,6]**. Повтор вернул ALREADY_CLOSED_TEST, остатки[9,6], ровно1 COMMITTED. На отдельном fixture controlled `during_publish` worker exit73 дал ожидаемый FAIL; restore вернул RECOVERED_ABORTED, decrements0 и **[12,10]**, exact initial XLSM SHA и все **11 VBA/OLE streams** равны. Это transaction crash, не дополнительный lifecycle Office-alive crash. DEV9/2026 close-preview вернул ожидаемый PASSPORT_9_2026 STOP до изменения остатков. Общая17-case negative suite, включая отдельный official=true случай, NOT_RUN.

## 9. WORD / EXCEL INTEGRITY

**PASS в указанном native/structural scope.**6/7 на комплект, исходные templates и sealed DEV data неизменны. Качество: point12 XML, constant cells, numbering, mergefields, tables/sections, opaque parts, points13/14 и TEST label validation PASS. Визуально осмотрены только **2 свежих native quality pages первого9/2026**: TEST-запрет официальной выдачи, масса, даты, оба ВИК/УЗК/АКЗ, пункты1–14 читаемы без замеченного clipping, ручная подпись пустая. Остальные страницы не объявлены визуально проверенными. {evidence('VISUAL_WORD_REVIEW.json')}.

## 10. PRODUCTION ISOLATION

**PASS в scope review по path contracts, фактическим26 JOB_REQUEST и before/after hashes.** Все mutating paths — новые parent/app-data/review roots; source только sealed DEV. Bootstrap явно читает source/DEV/ВЗМК_DEV.xlsx, bind/core направляют writers в owned workspace/run/artificial sandbox. Рабочий ВЗМК.xlsx не открывался для E2E, не записывался, не восстанавливался и не заменялся; постоянство его SHA не требовалось. Исторический writer UNKNOWN. Real deductions /official closures /deployment не выполнялись. Полного системного file-I/O trace нет; отсутствие операций записи подтверждено в проверенном пути, не как глобальное наблюдение всех программ пользователя. {evidence('PRODUCTION_ISOLATION_AUDIT.json')}.

## 11. SOURCE / PRODUCT INTEGRITY

**PASS. 67 804/67 804 защищённых файла byte-identical**, changed0 /added0 /deleted0 вне нового review evidence и index cache. Все **56 accepted Stage1–8 code files**, sealed DEV files,0.9.9 source/EXE/onedir, старая0.9.7 onedir и принятые outputs сохранились.0.9.7 не запускалась. Work не исправлял код и не пересобирал EXE. Пять существовавших Stage9 изменений к0.9.7 зафиксированы; их hashes до/после совпадают. {evidence('FINAL_IMMUTABLE_INTEGRITY.json')}, {evidence('SOURCE_INTEGRITY_AFTER.json')}, {evidence('SOURCE_SHA256_AFTER.json')}, {evidence('CHANGED_PRODUCT_FILES_AFTER.json')}.

## 12. FINDINGS

**BLOCKER0 /MAJOR0 /MINOR0 новых подтверждённых продуктовых дефектов.** Прежние fixes подтверждены source/compiled-проверками; обязательная независимая runtime GUI приёмка остаётся открытой. UNKNOWN exits сохранены как safety observations, без выдуманного дефектного файла/строки или минимального исправления. {evidence('CODEX09_FINDINGS.md')}.

## 13. NOT_RUN

Настоящий GUI E2E, viewer reopen/switch/first shutdown, frozen GUI Word.Close/no false SUCCESS; дополнительный lifecycle error/cleanup-error/Office-alive crash/restart; collision/repeated-native protocol;17-case negative suite; overlapping probes/controlled foreign/forged ownership. Дополнительно: все страницы визуально, АКЗ1/2/3, physical offline, полный OS I/O trace. Unknown user Office continuity не подтверждена. Подробные границы и причины — {evidence('NOT_RUN.json')}; короткий оставшийся сценарий — {evidence('USER_E2E.md')}.

## 14. FINAL VERDICT / READINESS

**FINAL VERDICT: FAIL. READY FOR USER E2E: NO. READY FOR PRODUCTION: NO.** Причина — незакрытые обязательные независимые проверки после STOP; полный PASS и отсутствие регрессий не подтверждены. Это не утверждение о найденном новом MAJOR/BLOCKER.

## 15. EVIDENCE / STOP

Полный {evidence('EVIDENCE_MANIFEST.json')} создаётся последним эксклюзивно: SHA собственных scripts/logs/results, новых external DEV/TEST roots, protected file maps и hashes исторических handoff references. {evidence('SEAL_VERIFY.json')} фиксирует pre-seal проверку; post-write manifest seal/readback и rehash выполняются seal-контроллером и выводятся в его результат без дописывания файлов после manifest. Authoritative controller hashes — {evidence('CONTROLLER_SHA256_SEALED.json')}. Raw FAIL, UNKNOWN и NOT_RUN сохранены. {evidence('REVIEW_RESULT.json')}.

**STOP. Product не изменён. Production deployment, реальные списания и официальные закрытия не выполнялись.**
''')
print(json.dumps(result,ensure_ascii=False,indent=2))
