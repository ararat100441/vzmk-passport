"""Write final self-test evidence and gated Work assignment; no GUI operations."""
from common import *
import shutil,win32api
from stage9_app.ownership import office_image
from stage8_closure.storage import digest
required=('SOURCE_TARGETED_17.json','VIEWER_REPEAT_SOURCE.json','WINDOW_CLOSE_GUARD_SOURCE.json',
          'WINDOW_CLOSE_EXTENDED_SOURCE.json','LIFECYCLE_FINAL.json','EXE_E2E_FINAL.json','EXE_NEGATIVE_FINAL.json',
          'REPEATED_NATIVE_OPEN_FINAL.json','OFFICE_DIAGNOSTIC_COLLISION_FINAL.json','NORMAL_OVERLAP_FINAL.json',
          'OVERLAP_CRASH_FINAL.json','FOREIGN_OFFICE_FINAL.json','OWNERSHIP_STOP_FINAL.json',
          'DOCUMENT_AND_XLSM_AUDIT.json','MANIFEST_AND_TRANSACTION_RECORDS_AUDIT.json','FINAL_PROCESS_AUDIT.json')
assert all(read(OUT/name)['status']=='PASS' for name in required)
before=read(OUT/'BUILD_SHA256_BEFORE.json');after=read(OUT/'BUILD_SHA256_AFTER.json');assert before==after
src=read(OUT/'SOURCE_SHA256_AFTER.json');assert src==read(OUT/'SOURCE_SHA256_BEFORE.json')
changed=read(OUT/'CHANGED_PRODUCT_FILES_AFTER.json');assert changed==read(OUT/'CHANGED_PRODUCT_FILES_BEFORE.json')
assert shutil.which('python',path=ENV['PATH']) is None
versions={}
for progid in ('Excel.Application','Word.Application'):
    path=office_image(progid);v=win32api.GetFileVersionInfo(path,'\\');ms=v['FileVersionMS'];ls=v['FileVersionLS']
    versions[progid]={'path':path,'file_version':'.'.join(map(str,(ms>>16,ms&65535,ls>>16,ls&65535)))}
save(OUT/'RUNTIME_ENVIRONMENT.json',{'exe_sha256':sha(EXE),'PATH':ENV['PATH'],'python_in_PATH':None,
    'PYTHONHOME_present':bool(ENV.get('PYTHONHOME')),'PYTHONPATH_present':bool(ENV.get('PYTHONPATH')),
    'Office':versions,'GUI_permission':False,'GUI_smoke':'NOT_RUN_USER_STOP','physical_offline':'NOT_RUN'})
controllers={str(p):sha(p) for p in OUT.glob('*.py')};controllers[str(ROOT/'validation/resume099_prepare.py')]=sha(ROOT/'validation/resume099_prepare.py')
save(OUT/'CONTROLLER_SHA256_FINAL.json',{'files':controllers,'aggregate_sha256':digest(controllers),
    'changes_from_prior_controllers':['fresh roots and result paths','both Tk smoke calls omitted','foreign sentinels explicitly hidden','self-test labels','read-only process exit race handled before image query'],
    'product_files_changed_this_resume':[]})
audit=read(OUT/'FINAL_PROCESS_AUDIT.json');e=read(OUT/'EXE_E2E_FINAL.json');detail=read(e['evidence'])
assert len(detail['steps'])==26 and detail['gui_smoke']=='NOT_RUN_USER_STOP'
def link(name):return '['+name+'](<'+(OUT/name).as_posix()+'>)'
table='\n'.join('| '+Path(r['path']).name+' | `'+r['sha256']+'` |' for r in changed)
report=f'''**ВЗМК — продолжение 0.9.9 после пользовательской остановки**

12.10.2026, Europe/Samara. Самопроверка Codex. Независимый Work не запускался. Computer Use не инициализировался/не возобновлялся; GUI не управлялся. Сохранённое окно EXE PID 47440 оставлено без управления. Все новые записи — только новые DEV/TEST roots из SETUP.json.

1. **Текущая версия и исходники.** Сохранена существующая **0.9.9**. Новых изменений продуктового кода в этом продолжении: **0**. Уже внесённые исправления не повторялись, новая сборка не требовалась. Дельта к reviewed 0.9.7 — пять файлов ниже; точные before/after SHA в {link('CHANGED_PRODUCT_FILES_AFTER.json')}. Source aggregate текущих runtime/spec/test: `{src['aggregate_sha256']}`; полная карта — {link('SOURCE_SHA256_AFTER.json')}.

| Уже изменённый файл | Текущий SHA-256 |
|---|---|
{table}

Диагностика остаётся exclusive/append-only: уникальный viewer directory, полные UUID событий, один финальный ownership.json после emergency outcomes и unconditional handle.Close. Word.Close ошибки сохраняются структурированно (step/type/HRESULT/excepinfo/traceback); Excel cleanup независим; GUI caller получает FAIL и не объявляет успешное открытие/закрытие. Источник: сохранённые исправления 0.9.9, не новая правка этого продолжения.

2. **Соответствие EXE и полнота onedir — PASS.** EXE: [{EXE.name}](<{EXE.as_posix()}>). SHA-256 **`{sha(EXE)}`**. Все **1117 файлов** совпали с сохранённым NEW_EXE_BINDING.json до/после тестов; aggregate `{after['aggregate_sha256']}`. **49 compiled modules / 60 raw source copies** соответствуют текущим исходникам; нормализован только co_filename. Весь прежний SOURCE_BEFORE_TESTS.json также совпал. Карты: {link('BUILD_SHA256_AFTER.json')}, {link('COMPILED_SOURCE_AFTER.json')}; зависимости — {link('DEPENDENCIES_AFTER.json')}. Для запуска нужен весь каталог onedir с _internal.

3. **Targeted source tests — PASS.** Повторены исходные **17/17** tests/stage9_098/test_cleanup.py, включая FileExistsError/exclusive evidence, disk/log errors, Word.Close/HRESULT, primary preservation, handle release, force ownership/durable begin и безопасный STOP. Дополнительно: 8 основных + 7 расширенных ownership/window races; три mock viewer reopen/close цикла с разными diagnostics, dedup и повторным idempotent close. Это source tests, а не frozen GUI PASS. GUI callbacks выполнялись только на fake widgets; реальный Tk/GUI не создавался. Evidence: {link('SOURCE_TARGETED_17.json')}, {link('VIEWER_REPEAT_SOURCE.json')}, WINDOW_CLOSE_GUARD_SOURCE.json, WINDOW_CLOSE_EXTENDED_SOURCE.json.

4. **Targeted EXE tests — PASS.** Саму EXE запустили для normal/error/cleanup-error/crash/normal-after-crash lifecycle; обе cleanup попытки продолжаются, primary error сохраняется. Четыре дополнительных native-verify цикла одного/другого/снова первого комплекта: реальные Word/Excel read-only opens, SHA документов неизменны, новые evidence paths, FileExistsError в обычном повторе отсутствует. Это CLI native verifier, НЕ GUI Viewer. Отдельно две реальные fault-пробы EXE: контроллер резервировал office.json только своего доказанного job до final save; программа вернула ожидаемый FAIL/STOP, не перезаписала evidence, оба Office вышли 0 без force. При сочетании collision + cleanup-error сохранены STAGE9_TEST_OFFICE_ERROR, Word.Close error и finalization diagnostic. FileExistsError здесь намеренная отрицательная проба, не скрытый normal FAIL. Evidence: {link('REPEATED_NATIVE_OPEN_FINAL.json')}, {link('OFFICE_DIAGNOSTIC_COLLISION_FINAL.json')}, LIFECYCLE_FINAL.json.

5. **EXE CLI integration/E2E — PASS в доступной области.** **26 actual frozen tasks**, без двух GUI-producing smoke шагов. Три полных TEST-комплекта: 9/2026, новый 10/2026, повторный 9/2026; Приложение №1, два ВИК в одном DOCX, УЗК SAMPLE_ONLY, выбранный АКЗ 0 и качество последним. Каждый: 104 visible / 397 hidden excluded / 288 шт. / 15822.14 кг / 1913.494 м² / УЗК 33; M1/connections, numbers/dates, Word 3 / Excel 3 проверены. Итого **18 files / 21 records**. Далее official TEST closure STOP; искусственный 8001/2099 CLOSED_TEST, остатки [12,10]→[9,6]; repeat ALREADY_CLOSED_TEST без списания; journal/recovery; crash during_publish; restore RECOVERED_ABORTED с decrements=0, остатками [12,10], exact seed XLSM SHA и всеми **11 VBA/OLE streams**. Manifest seals, input hashes, native evidence hashes и journal hash chains проверены. {link('EXE_E2E_FINAL.json')}, {link('DOCUMENT_AND_XLSM_AUDIT.json')}, {link('MANIFEST_AND_TRANSACTION_RECORDS_AUDIT.json')}.

Негативные EXE проверки — **17/17 PASS**: missing/invalid VOM, missing/tampered template, long Word path STOP до Office, preview/source tamper, incomplete manifest, official flag/9, missing UZK, incomplete AKZ, unconfirmed control, Windows lock, before_save/save_io, damaged backup STOP. {link('EXE_NEGATIVE_FINAL.json')}. Шесть overlapping normal, normal рядом с crash, две normal рядом с pumping controlled foreign Office и forged ownership STOP прошли. Реальные контрольные foreign Word/Excel сохранили identity и содержимое; source-controller затем завершил только свои sentinels. Новые процессы продукта: **{audit['job_count']} jobs / {audit['office_instances']} Office instances**, normal force **0**, own orphans **0**, unknown child identities **0**, validated WM_CLOSE events **{len(audit['WM_CLOSE_events'])}**. Аварийных force **{len(audit['force_proofs'])}**, только intentional worker exit 73, original CREATE_PROCESS handles/identity/FORCE_BEGIN проверены. Token elevation 0. {link('FINAL_PROCESS_AUDIT.json')}.

Все EXE-команды использовали PATH только System32/Windows, PYTHONHOME/PYTHONPATH удалены, python в PATH отсутствует; worker запускается тем же EXE. {link('RUNTIME_ENVIRONMENT.json')}. Python source tests не выдаются за EXE tests. Физическое отключение сети **NOT_RUN**; приложение сохраняет существующий no_network guard. Это не объявлено physical Office offline PASS.

6. **Source integrity — PASS.** Все **56 accepted Stage 1–8 code files** unchanged; вся 0.9.7 onedir unchanged; существующая 0.9.9 и current source до/после byte-identical. Sealed DEV source unchanged. Реальные реестры, ВОМ, справочники, готовые паспорта не изменялись и для E2E не использовались. Рабочий ВЗМК.xlsx не включён в SHA gate; его исторический writer остаётся UNKNOWN. Workspace этого запуска: `{e['workspace']}`. {link('SOURCE_INTEGRITY_AFTER.json')}. Добавлены только validation controllers/evidence; их hashes/provenance — {link('CONTROLLER_SHA256_FINAL.json')}, CONTROLLER_PROVENANCE.json. Начальный process audit встретил завершавшийся baseline process (WinError 31 при image query); read-only контроллер исправлен на проверку exit перед image query. Продукт не менялся; первоначальный отказ сохранён в PROCESS_AUDIT.log, успешный аудит — PROCESS_AUDIT_R2.log. Baseline Word 3132 к концу не жив: причина UNKNOWN, непрерывность этой реальной сессии NOT_VERIFIED; он не относится к новым child records/force targets. GUI PID 47440 сохранил identity и не управлялся. Контролируемая foreign Office проверка отдельно PASS.

7. **GUI и готовность.** **GUI-click E2E: NOT_RUN_USER_STOP**. Настоящие frozen GUI Viewer reopen/switch/first shutdown и GUI propagation Word.Close ещё не приняты: source mocks/CLI verifier их не подменяют. Причина — прямой запрет пользователя после Escape, а не недоступность инструмента. Остался обязательный GUI targeted этап, требующий отдельного разрешения. Исторические transient Add.Content/Office retention observations не переименованы в доказанный root cause; новых unexpected runtime FAIL в этой CLI серии не обнаружено.

**READY FOR WORK 09 REVIEW: NO** — полный targeted кандидат ожидает разрешённой GUI самопроверки замечания W09-097-M01. Задание и headless evidence готовы; независимый Work не запускался. **READY FOR USER E2E: NO. READY FOR PRODUCTION: NO.** Production deployment, official closure и реальные списания не выполнялись. После отчёта STOP и запрос разрешения только на оставшуюся GUI проверку.

Evidence manifest формируется последним: EVIDENCE_MANIFEST.json, EVIDENCE_MANIFEST_SHA256.json, SEAL_VERIFY.json. GUI NOT_RUN сохранён в manifest, независимо от остальных PASS.
'''
with (OUT/'STAGE9_REPORT.md').open('x',encoding='utf8') as f:f.write(report)
assignment=f'''**WORK 09 targeted assignment — existing 0.9.9, gated by user GUI permission**

Подготовлено; НЕ запущено. Readiness release: NO до оставшейся GUI самопроверки. Все переданные PASS — самопроверка Codex, не независимый Work review.

ROOT: {ROOT}
EXE: {EXE}
SHA-256: {sha(EXE)}
Onedir: 1117 files, aggregate {after['aggregate_sha256']}
Source aggregate: {src['aggregate_sha256']}

Прочитать STAGE9_REPORT.md, полный EVIDENCE_MANIFEST/SEAL_VERIFY, BUILD_SHA256_AFTER, COMPILED_SOURCE_AFTER, SOURCE_SHA256_AFTER, CHANGED_PRODUCT_FILES_AFTER. Сверить текущий EXE/все 1117 файлов, 49 compiled/60 raw и 56 accepted Stage 1–8 hashes. Проверить исходные замечания validation/work09_targeted_097_20261012_01/CODEX09_FINDINGS.md: duplicate final diagnostic/exclusive FileExistsError и discarded viewer Word.Close errors. Не переносить результаты 0.9.7 в PASS 0.9.9.

Прямое ограничение пользователя: Computer Use остановлен Escape; НЕ возобновлять Computer Use, не управлять/не запускать GUI без нового явного разрешения. Не закрывать сохранённое окно PID 47440. До разрешения допустимы только source read-only и hidden EXE CLI проверки на новых своих TEST-копиях; GUI, включая --task smoke (создаёт Tk), обозначать NOT_RUN. Не обходить запрет через другой UI-инструмент, PowerShell UI Automation или передачу GUI-действий другому агенту.

1. Проверить single final ownership write после emergency cleanup, unconditional handle.Close/STA restoration, сохранение primary exception. Сохранить exclusive creation, не заменять overwrite. Проверить уникальные viewer attempt directories/events, reserved-name retry, idempotent close. Source 17 tests + 15 guard cases + 3 mock viewer cycles переданы, но Work получает собственные результаты. Диагностический disk/collision failure обязан давать FAIL/STOP, old evidence unchanged; force без proven ownership/durable begin запрещён.

2. Проверить Viewer.close Word.Close: structured step/type/HRESULT/excepinfo/traceback, Excel.Close/Quit попытаны, primary opening exception сохранена. GUI close_viewers обязан показать FAIL и не запускать новый viewer/не объявлять успешный shutdown. Frozen office-probe cleanup-error относится к entry-point probe; это не native fault injection GUI Viewer. Source mocks также не GUI PASS.

3. Самостоятельно повторить hidden frozen normal/error/cleanup-error/crash/restart, overlapping workers, controlled foreign Office и forged ownership. Только actual child/PID/birth/image/original CREATE_PROCESS handle; normal force всегда FAIL. Проверить real exit, own orphan 0; не завершать неизвестный/пользовательский процесс. Baseline пользователь может закрыть сам: сохранить UNKNOWN, если автор завершения не доказан. Переданный baseline Word 3132 ended, cause UNKNOWN; controlled foreign Office PASS.

4. Для CLI E2E создать новый короткий TEST parent и новый LOCALAPPDATA/TEMP. Source только sealed DEV `{SOURCE}`, ВОМ `{VOM}`; источник read-only, никаких production данных/изменений. Создать собственный workspace через новую EXE --task import. Не запускать переданные mutating controllers поверх текущих FINAL файлов: скопировать/адаптировать контроллеры только к новым outputs/roots. Передан 26-task CLI E2E (GUI smokes исключены), 3 комплекта 18 files/21 records, native Word/Excel read-only, M1, official STOP, artificial 8001/2099 once-only close, crash/restore exact XLSM SHA/VBA 11 streams. Work обязан получить свои результаты.

5. Повторить 17 safe negatives: input/template/SHA/preview/manifest conflicts, long Word pending path before COM, official/unconfirmed/incomplete UZK/AKZ, Windows lock, save faults/recovery и damaged backup STOP. Только искусственные fixtures. Сохранить Stage 8 backups/locks/journal/idempotency/recovery без правок accepted code.

6. Повторить четыре native CLI read-only opens одного/другого/снова первого набора, уникальные validation paths/SHA unchanged; отдельно actual EXE final diagnostic collision: резервировать ТОЛЬКО office.json своего доказанного supervisor/job, ожидать FAIL и сохранность old evidence. Collision + primary cleanup-error должен сохранять primary/Word.Close и secondary finalization failure. Эти два режима у Codex прошли, но не являются обычным GUI повторным открытием.

7. ТОЛЬКО после нового явного разрешения пользователя: сама EXE 0.9.9 GUI, DEV project/passport/VOM; два последовательных Open Word/Excel, switch bundle, first shutdown с открытым viewer. Подтвердить absence FileExistsError, новые diagnostics, 3 read-only DOCX/3 XLSX, unchanged document SHA, оба logical ВИК, Word.Close error surfaced/durable и safe STOP, foreign Office untouched, no own orphans. Сделать собственное GUI E2E; не считать прежнюю остановку/первое окно/CLI mocks PASS. Отдельное разрешение и результаты должны войти в новое evidence.

Границы: Stage 1–8 READ-ONLY; реальные реестры/ВОМ/справочники/готовые паспорта не менять/не использовать для E2E; никакого official closure, реального УЗК списания, production deployment, global taskkill/UAC/global Office addin changes. Рабочий ВЗМК.xlsx не восстанавливать/заменять/синхронизировать; writer UNKNOWN и production SHA не gate при параллельной работе пользователя. При доказанной production записи продукта STOP с exact operation/path.

Вернуть независимые findings/verdict, SHA до/после, source/onedir/Office/CLI/GUI/negative results, все FAIL/NOT_RUN и readiness. Код продукта не исправлять в review. После отчёта STOP. Если GUI запрещён/не проверен — полный targeted review/User E2E не объявлять PASS/ready.
'''
with (OUT/'WORK09_TARGETED_ASSIGNMENT.md').open('x',encoding='utf8') as f:f.write(assignment)
save(OUT/'REVIEW_STATUS.json',{'version':'0.9.9','exe_sha256':sha(EXE),'headless_self_tests':'PASS','assignment_prepared':True,'independent_review':'NOT_RUN','GUI':'NOT_RUN_USER_STOP','ready_for_work09_review':False,'ready_for_user_e2e':False,'ready_for_production':False,'remaining_required_gate':'Explicit permission and actual frozen GUI viewer reopen/switch/shutdown validation','product_changed_this_resume':False})
print('Report and Work assignment prepared; GUI permission gate remains; no product edits')
