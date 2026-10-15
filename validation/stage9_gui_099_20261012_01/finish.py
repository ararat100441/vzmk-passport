"""Finalize explicit GUI scope and a prepared, unstarted independent assignment."""
from prepare import *
import gc
v=read(OUT/'GUI_VALIDATION.json');assert v['status']=='PASS'
# Verify only the external controller's fixed Popen destructor bookkeeping.
child=subprocess.Popen([sys.executable,'-B','-c','pass'],stdout=subprocess.PIPE,stderr=subprocess.PIPE,creationflags=subprocess.CREATE_NO_WINDOW)
win32event.WaitForSingleObject(child._handle,10000);code=win32process.GetExitCodeProcess(child._handle)
assert code==0;child.returncode=code;child._handle.Close();del child;gc.collect()
save(OUT/'CONTROLLER_HANDLE_BOOKKEEPING_CHECK.json',{'status':'PASS','scope':'VALIDATION_CONTROLLER_ONLY_NOT_EXE_TEST','product_changed':False,'child_exit_code':code})
current=(OUT/'watch_gui.py').read_text(encoding='utf8')
old=current.replace("exit_code=win32process.GetExitCodeProcess(h)\nsave(OUT/('GUI_'+mode+'_EXIT.json'),{'pid':pid,'creation_time':birth,'exit_code':exit_code,'exit_observed':True,'forced':False})\nif mode!='EXISTING':child.returncode=exit_code  # Prevent Popen.__del__ querying an already closed handle.\n", "save(OUT/('GUI_'+mode+'_EXIT.json'),{'pid':pid,'creation_time':birth,'exit_code':win32process.GetExitCodeProcess(h),'exit_observed':True,'forced':False})\n")
assert current!=old
save(OUT/'CONTROLLER_OBSERVATIONS.json',{
 'warning_observed':'External validation controller watch_gui.py FRESH: after GUI EXIT PASS, Popen.__del__ emitted OSError [WinError 6] invalid handle.',
 'cause':'Controller manually closed retained child handle before setting Popen.returncode; destructor tried querying that closed handle.',
 'product_effect':'None: frozen GUI exit code 0 was already recorded through valid original handle; controller exit code 0; no product mutation.',
 'fix':'Set child.returncode before closing the handle; controller-only check PASS.',
 'executed_watch_gui_sha256':hashlib.sha256(old.encode('utf8')).hexdigest(),
 'final_watch_gui_sha256':sha(OUT/'watch_gui.py'),
 'ui_tool_observations':['Initial window minimized: activated only returned EXE window.', 'System file dialog accessibility set_value refused cached element; fresh screenshot click/type succeeded.', 'One occluded capture showed another window; no input sent from that capture; reselected and activated EXE.', 'GUI_16_RECOVERY_OPEN_PASS is a prematurely named intermediate screenshot still showing the expected error state, NOT a PASS assertion. GUI_17_RECOVERY_CONFIRMED and sixth OPENED event prove eventual explicit reopen.', 'Immediate window list after Alt+F4 still contained closing window; final confirmed list is empty.'],
 'product_defects_found':[]})
report=f'''**ВЗМК — Stage 9 / EXE 0.9.9 / ограниченный DEV GUI E2E**

12.10.2026, Europe/Samara. Самопроверка Codex по новому явному разрешению пользователя на ограниченный Computer Use. Все действия интерфейса — skill computer-use / node_repl @oai/sky. Независимый Work НЕ запускался. Прежний sealed отчёт сохранён как исторический снимок; этот отчёт заменяет только его GUI NOT_RUN/readiness gate новыми фактическими результатами.

**GUI E2E: PASS в разрешённой области.** Шесть реальных viewer open/close циклов собранной EXE, без source monkeypatch/in-process test harness. Одно ожидаемое прикладное FAIL при отрицательной пробе явно сохранено.

| Проверка | Результат |
|---|---|
| Главное окно / свежий запуск той же EXE | PASS: сохранённый PID 47440 и свежий PID 49108, точные birth/image/SHA |
| TEST Word и Excel | PASS: каждый цикл 3 DOCX + 3 XLSX; UI read-only; документы и весь TEST-стенд byte-identical после тестов |
| Два повторных открытия того же набора | PASS: циклы 1→2→3; прежний viewer закрывается, создаются новые диагностики |
| Несколько последовательных циклов | PASS: 6 opens / 6 closes; 5 обычных close PASS, 1 ожидаемый close FAIL |
| Смена TEST-комплекта | PASS: 9/2026 → 10/2026 → тот же 9/2026; OPEN_BEGIN пути совпадают с точными manifests |
| Повторный FileExistsError | PASS: 0; шесть уникальных viewer directories/attempt IDs, один final ownership.json на viewer |
| Word.Close ошибка / safe STOP | PASS отрицательной проверки: после GUI-закрытия только собственного Word PID 51692 exit 0 вызван повторный просмотр; реальный Word.Close com_error HRESULT -2147023174 / RPC unavailable; GUI показывает Word.Close, Word.Quit, путь диагностики и «Новый просмотр не запущен». Нового viewer в этой попытке нет, Excel закрыт независимо. После подтверждения и нового явного клика шестой просмотр проходит |
| Штатный выход EXE | PASS: одна команда Alt+F4 с активным viewer, PID 47440 exit 0; свежий PID 49108 также закрыт с первой команды, exit 0 |
| Office orphan / force | PASS: 12/12 собственных Office завершились exit 0; orphan 0, forced termination 0, FileExistsError 0 |

Процессная доказательность: перед GUI Office baseline **0 Word / 0 Excel**. Каждый новый Office связан по PID, creation time, image, parent=47440 и original CREATE_PROCESS ownership record. Независимый read-only монитор удерживал QUERY_LIMITED_INFORMATION|SYNCHRONIZE handles, отслеживал появляющиеся Office и их exit; API завершения в контроллере нет. Неожиданных non-test Office exit: **0**; неизвестные процессы не завершались. В этой GUI серии не было активных пользовательских Office, поэтому она не выдаётся за тест сохранности существующей пользовательской Office-сессии. Отдельные controlled foreign Office проверки предыдущей CLI серии остаются PASS. Исторический Word **PID 3132: причина UNKNOWN**, этот запуск её не установил.

Excel после Quit в шести циклах потребовал существующую штатную owned-empty WM_CLOSE процедуру: PID/birth/image/original handle, пустые Workbooks и COM refs=0 проверены продуктом; каждый Excel вышел 0. WM_CLOSE события сохранены. TerminateProcess в этой GUI серии **не вызывался**.

Версия **0.9.9** сохранена, пересборка не нужна: новых дефектов продуктового Stage 9 в этой серии не найдено; продуктовых изменений **0**. EXE: `{read(OUT/'BINDING_AFTER.json')['build']['exe']}`. SHA-256: **{v['exe_sha256']}**. Все **1117 onedir files** совпали до/после; aggregate `0a9abec9d2583aaf4370eb9aa143d9bed63ebd7561b50b4df497ce84f29cad01`. Current source aggregate `913ce2bc334ed2dbfae70311089000bf0b47220b5caef3b591e4175b032c983c`; **56 accepted Stage 1–8 hashes** unchanged. Предыдущая проверка 49 compiled/60 raw остаётся привязанной к тем же неизменным source/EXE hashes. Ранее изменённые пять Stage 9 файлов и точные SHA — CHANGED_PRODUCT_FILES_AFTER.json прежнего sealed набора; здесь они не редактировались.

Использован только ранее созданный изолированный TEST-стенд `{read(OUT/'BINDING_AFTER.json')['workspace']}` и его два manifest. Реальные реестры ВЗМК/УЗК, ВОМ, справочники, готовые паспорта не использовались и не изменялись. Генерация/официальное закрытие/списание в этой ограниченной GUI серии не запускались. Прежние EXE CLI 26 задач, 17 negative tests, source 17/17 сохранены отдельно; их полные 2160 sealed files повторно сверены, mismatch 0. GUI PASS не подменяет полноценную GUI генерацию, которую этот запрос не разрешал. Свежий GUI запущен контроллером с изолированными LOCALAPPDATA/TEMP и PATH только System32/Windows; далее управляется исключительно sky, внешнего Python в PATH нет.

Обнаруженные ошибки: ожидаемый Word.Close/RPC FAIL сохранён полностью; неожиданных product FAIL нет. Сбои accessibility/промежуточный не сработавший click не считались PASS, повтор выполнялся только после нового наблюдения. У внешнего контроллера после доказанного выхода свежей GUI была destructor warning WinError 6 из-за ручного Close handle без Popen.returncode; исправлен только validation/watch_gui.py, отдельная controller-only проверка PASS. Подробности, включая ошибочно раннее имя GUI_16, сохранены в CONTROLLER_OBSERVATIONS.json. Это не ошибка EXE и не основание менять её SHA.

Добавлены только validation scripts, screenshots, metadata и копии собственных viewer diagnostics в этом новом каталоге. BINDING_BEFORE/AFTER, GUI_VALIDATION, OFFICE_BASELINE/MONITOR, original diagnostic binding, GUI error accessibility/screenshot, GUI exit records составляют доказательства. Evidence manifest/seal формируются последними. Старый GUI_NOT_RUN отчёт не перезаписан.

**READY FOR WORK 09 REVIEW: YES.** Кандидат и targeted assignment готовы; независимый verdict **NOT_RUN**. **READY FOR USER E2E: NO** до независимого Work 09 и необходимых acceptance checks. **READY FOR PRODUCTION: NO**. Work самостоятельно не запускался, deployment не выполнялся. После отчёта **STOP**.
'''
(OUT/'STAGE9_GUI_REPORT.md').write_text(report,encoding='utf8')
assignment=f'''**WORK 09 TARGETED ASSIGNMENT — 0.9.9 after completed limited DEV GUI self-test**

PREPARED ONLY. НЕ ЗАПУЩЕНО. READY FOR WORK 09 REVIEW: YES означает готовность кандидата, не независимый PASS и не разрешение автоматического запуска. Пользователь прямо запретил запускать Work самостоятельно. Новое разрешение Computer Use распространялось исключительно на ограниченный Codex DEV GUI E2E этой серии; не считать его разрешением будущему Work на GUI без явного согласования.

EXE: {read(OUT/'BINDING_AFTER.json')['build']['exe']}
SHA-256: {v['exe_sha256']}
Onedir: 1117 files; aggregate 0a9abec9d2583aaf4370eb9aa143d9bed63ebd7561b50b4df497ce84f29cad01.
Source aggregate: 913ce2bc334ed2dbfae70311089000bf0b47220b5caef3b591e4175b032c983c.

Прочитать STAGE9_GUI_REPORT.md / GUI_VALIDATION.json / EVIDENCE_MANIFEST.json / SEAL_VERIFY.json этого каталога. Затем прежний sealed validation/stage9_resume_099_20261012_01: STAGE9_REPORT, full EVIDENCE_MANIFEST, SOURCE/BUILD/COMPILED_SOURCE_AFTER, CHANGED_PRODUCT_FILES_AFTER, SOURCE_INTEGRITY_AFTER, source17, actual EXE26 E2E / negative17 / collision / repeated native / foreign Office evidence. Прежний GUI NOT_RUN — исторический статус до разрешения, не скрывать его. Сверить все hashes; self-test Codex не выдавать за независимый Work.

Проверить исходные независимые findings validation/work09_targeted_097_20261012_01/CODEX09_FINDINGS.md: double exclusive final diagnostics и discarded Word.Close errors. В 0.9.9 single final ownership write после cleanup, unconditional handle release, unique UUID viewer/event directories и fail-closed cleanup propagation должны сохраниться. Не менять exclusive creation на overwrite и не отключать ownership/durable evidence checks.

GUI self-evidence: шесть настоящих frozen viewer циклов 9,9,9,10,9,9; 3 Word/3 Excel read-only в каждом, exact document/TEST-workspace SHA unchanged; пятый close ожидаемо FAIL после ручного Alt+F4 закрытия трёх собственных Word окон, Word PID 51692 вышел 0. Следующий Open blocked, показан Word.Close HRESULT -2147023174 / RPC unavailable и diagnostic path; Excel cleanup независим и exit 0. Новый явный Open после acknowledgement проходит; шестой viewer закрыт при первой команде выхода EXE, exit 0. Свежий GUI launch/first exit тоже pass. 12 own Office, force0, orphan0, foreign baseline0. Проверить JSON/screenshot фактическое содержание; GUI_16 имя преждевременно, PASS подтверждается GUI_17 и OPENED6, а не именем файла.

После отдельного разрешённого запуска независимого Work повторить targeted source17 + реальный frozen normal/error/cleanup-error/crash/restart и native reopens, safe ownership / diagnostic collision failure. При согласованном GUI разрешении повторить настоящие viewer reopen/switch/first shutdown и visible Word.Close safe STOP на собственном свежем TEST workspace. Нельзя подменять actual frozen GUI source mocks или CLI verifier. Не запускать mutating controllers поверх переданных FINAL файлов; создать новые output/LOCALAPPDATA/TEMP/short workspace, источник только sealed DEV C:\\Users\\arara\\AppData\\Local\\VZMKStage9_092_9c0a3396\\source, TEST_VOM.xls. Stage8 регрессии backup/locks/journal/once-only/recovery проверять только на новых artificial fixtures; принятую логику не править.

Перед тестом baseline Office identity PID/birth/image; unexpected non-test Office exit → немедленно STOP и диагностика. PID 3132 historical ended, cause UNKNOWN. Контроллеры и продукт могут действовать только на доказанный собственный Office; unknown ownership → STOP. Normal force — finding, не PASS. Сохранить все FAIL/NOT_RUN и первичные traceback/HRESULT/diagnostic hashes.

Границы: Stage1–8 READ-ONLY, production registry/BOM/catalog/passports не изменять, реальных списаний/официального закрытия/UAC/global taskkill/deployment нет. Review без продуктовых правок. Вернуть собственный независимый verdict/findings + EXE/source/onedirSHA, actual CLI/GUI/negative/Office/source integrity и readiness. После отчёта STOP.
'''
(OUT/'WORK09_TARGETED_ASSIGNMENT.md').write_text(assignment,encoding='utf8')
save(OUT/'REVIEW_STATUS.json',{'gui_e2e':'PASS_LIMITED_DEV_FROZEN_SCOPE','ready_for_work09_review':True,'independent_review':'NOT_RUN_USER_PROHIBITED','ready_for_user_e2e':False,'ready_for_production':False,'product_files_changed_this_turn':0,'new_build_required':False,'stop':True})
print('REPORT / WORK ASSIGNMENT PREPARED; Work not dispatched',flush=True)
