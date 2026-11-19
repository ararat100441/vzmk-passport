"""Write the handoff only after the final immutable-build checks pass."""
from common import *
audit=read(OUT/'FINAL_AUDIT.json');assert audit['status']=='PASS'
build=read(OUT/'BUILD_SHA256_FINAL.json');source=read(OUT/'SOURCE_SHA256_FINAL.json')
binding=read(OUT/'COMPILED_SOURCE_FINAL.json');integrity=read(OUT/'SOURCE_INTEGRITY_FINAL.json')
old=read(OUT/'OLD_09_INTEGRITY_FINAL.json');deps=read(OUT/'DEPENDENCIES_FINAL.json')
e2e=read(read(OUT/'EXE_E2E_FINAL.json')['evidence']);negative=read(read(OUT/'EXE_NEGATIVE_FINAL.json')['evidence'])
life=read(read(OUT/'LIFECYCLE_FINAL.json')['evidence'])
assert build['exe_sha256']==sha(EXE)==e2e['exe_sha256']==negative['exe_sha256']==life['exe_sha256']
assert integrity['status']=='PASS' and binding['status']=='PASS'
previous=read(ROOT/'validation/stage9_20261012_01/SOURCE_SHA256_FINAL.json')['files']
changed={p:{'previous':previous.get(p),'current':h} for p,h in source['files'].items() if previous.get(p)!=h}
save(OUT/'CHANGED_FILES.json',{'stage':'Stage 9 only','files':changed,'additional_generated_artifacts':['stage9_dist_091/','stage9_build_091/','STAGE9_DEV_091.md','validation/stage9_091_20261012_01/'],'accepted_stage1_8_changed':False})
package_versions='; '.join(f'{k} {v}' for k,v in deps['packages'].items())
office_versions='; '.join(f"{k}: {v['version']}" for k,v in audit['office_versions'].items())
runtime_count=sum(r['included_in_pyz'] for r in binding['modules'])
code_changes='\n'.join('- `'+p+'`' for p in changed)
steps='\n'.join(f"| {i} | {s['task']} | {s['expected']} | `{Path(s['result']).name}` |" for i,s in enumerate(e2e['steps'],1))
cases='\n'.join('| '+c['name']+' | '+c['status']+' |' for c in negative['cases'])
report=f'''# Stage 9 — Office worker fix, DEV/TEST 0.9.1 / R7

Дата: 12.10.2026. Собственные проверки текущей EXE пройдены. Независимый Work 09: **NOT_RUN**.

READY FOR WORK 09 REVIEW: **YES**  
READY FOR USER E2E: **NO**  
READY FOR PRODUCTION: **NO**

## 1. ROOT CAUSE

В сохранённой сборке 0.9 точка отказа — `_internal/stage9_app/ownership.py:78`, `TerminateProcess(handle, 0)` в finalization после ожидания только 3000 ms (строка 75). Завершался сервер Office, а не Python worker. `DispatchEx` создавал COM-сервер с DCOM-parent PID 2000, а код считал его собственным дочерним процессом лишь по новому PID/creation time. Handle получался повторным `OpenProcess`, исходного handle создания не было.

Преждевременность force доказана диагностическим frozen-прогоном: Quit вернулся, число COM interfaces было 0, но штатный выход Excel занял 52.547 s; после увеличенного ожидания — exit 0. См. `GRACE_CHECK.json`, `GRACE_LOG_INDEX.json`. Происхождение не как actual child доказано `PROCESS_EXTENDED_STATE.json`.

Причина не сведена к отсутствию администратора. Исправленный `HANDLE_ACCESS_PROOF_R2.json` показывает действующий handle, granted access 0x101401, включая PROCESS_TERMINATE, и STILL_ACTIVE 259. Первоначальный HANDLE_ACCESS_PROOF.json содержит ошибочный размер структуры NtQueryObject и не используется как доказательство прав. SECURITY_SNAPSHOT*.json: одинаковый SID, medium IL, elevated=0, protection=0; поздний снимок также проверяет AppContainer. Точный нижележащий компонент Windows/Office, вернувший error 5 при прежнем non-child force, **UNKNOWN**. Нет доказательства необходимости UAC и нет утверждения, что Windows запрещает завершать любые non-child процессы.

Установленный дефект управления — преждевременный force внутри COM worker плюс неподтверждённое предположение о child ownership. Новый код устраняет этот путь. Force на исходном handle настоящего дочернего процесса доказан отдельно и в аварийном тесте текущей EXE. Документация Microsoft требует PROCESS_TERMINATE и ожидание завершения: [TerminateProcess](https://learn.microsoft.com/en-us/windows/win32/api/processthreadsapi/nf-processthreadsapi-terminateprocess).

## 2. EXACT CODE FIX

- `stage9_app/ownership.py:26`: собственный Popen/CreateProcess Office; Excel /x /automation, Word /w /m; macro-free probe. Сохраняются исходный handle, PID, creation time, image, launcher PID. Для frozen external Office временно очищается SetDllDirectory, затем восстанавливается.
- `ownership.py:48`: OBJID_NATIVEOM привязывает COM к окну точного PID; проверяются handle/birth/image и PID NativeOM/Application. Исходная IDispatch reference AccessibleObjectFromWindow освобождается после ObjectFromAddress, добавляющего свою QI reference. Подключения к пользовательскому активному Office нет.
- `ownership.py:119`: accepted DispatchEx адаптируется только на время работы Stage 9. Workbook/Document закрывают принятые engine-контексты; Quit логируется. STA сохраняется до завершения worker; оставшиеся COM-ссылки освобождаются. Worker сам не force-terminates Office.
- `stage9_app/supervisor.py:35`: родитель создаёт и сохраняет actual Office children, проверяет launch request/START и держит исходные handles вне worker. После смерти worker — адресный helper Close(False)/Quit и 120 s grace. Только затем возможен force с повторной проверкой actual child, launcher PID, handle source, GetProcessId, creation time и image. Durable FORCE_BEGIN сохраняется до вызова. Unknown ownership даёт STOP; неизвестный Office не завершается. taskkill отсутствует.
- `stage9.py`: parent/worker/helper используют ту же EXE; PASS публикуется после cleanup. Искусственные fault hooks ограничены TEST 8001/2099. Для minimal PATH добавляется существующий системный Windows PowerShell, требуемый принятым renderer. Python в PATH не добавляется.
- `stage9_app/office.py`: перед выходом viewer guard очищается последняя локальная COM-ссылка app.
- Версия 0.9.1 в GUI/конфигурации, отдельный spec/каталог. Build AST preflight прекращает сборку при SyntaxError.

Принятая логика Stage 1–8 не изменена. Точный diff: `STAGE9_DIFF_FINAL.patch`; индивидуальные SHA: `SOURCE_SHA256_FINAL.json`.

## 3. OFFICE WORKER VALIDATION

`LIFECYCLE_FINAL.json` → `{life['workspace']}`. Четыре прогона именно текущей EXE: normal → handled error → hard worker death (exit 73) → normal. Обе Office application в каждом прогоне — actual children с original CREATE_PROCESS handles.

Normal/error завершаются без force. Hard death остаётся транспортным FAIL, как и ожидается: Word завершается адресным Quit; у Excel после Quit/grace понадобился fallback на исходном child handle, успешно, без error 5. FORCE_BEGIN содержит validated handle, точные PID/birth/image/launcher. Это аварийный выход, не назван штатным PASS процесса; PASS относится к корректной реакции теста. Во всех прогонах orphan=0.

`OWNERSHIP_STOP_FINAL.json`: подложный START на пользовательский Excel, owned=false/actual_child=false, вызвал безопасный STOP без adoption/kill. `FINAL_AUDIT.json`: проверено {audit['owned_records_checked']} proven ownership records, собственных живых orphan=0. Baseline Excel PID 51208 и неизвестный pre-existing Word PID 77680 сохранили PID/creation time. Word 77680 не присваивается программе и не закрывался.

Ранние собственные остатки R1–R3 устранены адресно после точной identity-проверки; `legacy_thread_stop_b15fa27e/RESULT.json` — два уже пустых TEST Excel exit 0 через WM_QUIT собственного root thread. Ранний Word завершился ненормально C0000005, это исторический FAIL, а не штатный выход. Глобальное завершение Office не применялось.

## 4. NEW EXE PATH / SHA

EXE: `{EXE}`  
SHA256: `{build['exe_sha256']}`  
Onedir: {build['file_count']} files; aggregate `{build['aggregate_sha256']}`. Переносить всю папку с `_internal`.

Source/tests aggregate: `{source['aggregate_sha256']}`. Все индивидуальные SHA в BUILD_SHA256_FINAL.json / SOURCE_SHA256_FINAL.json. Компилированные runtime-модули {runtime_count}, raw source copies {len(binding['modules'])}: соответствуют исходникам; 9 Stage 9 modules обязательны и совпали. Неиспользуемые CLI-модули, включённые только как raw data, помечены included_in_pyz=false.

CPython 3.12.7 x64; {package_versions}. Системные зависимости: {office_versions}; Windows PowerShell 5.1. Build-only WMI fallback shim не включён в runtime.

Сборка не содержит пользовательских реестров/паспортов/ВОМ. Единственный DOCX внутри — стандартный пустой python-docx/templates/default.docx, SHA совпадает с библиотекой. Прежняя 0.9: все {old['files']} файлов совпали с baseline; EXE `{old['exe_sha256']}`, aggregate `{old['aggregate_sha256']}`.

## 5. EXE SMOKE TESTS

Текущий frozen=true, version=0.9.1, GUI создано, python_in_path=null — шаги 2 и 28 E2E. Весь набор запускался с входным PATH только Windows/System32 и Windows; системный PowerShell разрешается самим entrypoint. Python не удалялся из системы.

Исходный GUI: `gui_1297e541/GUI_LAYOUT.json`, 9/9 bounds PASS (720×520, 1000×760, 1360×900; масштабы 1, 1.5, 2), stale selection и async callbacks PASS. Это source GUI tests, не девять EXE UX-прогонов. Screenshot с PrintWindow имеет артефакты кешированного окна и не используется как полная визуальная приёмка EXE.

## 6. FULL DEV E2E RESULTS

`EXE_E2E_FINAL.json` → `e2e_2b9df014/RESULT.json`: PASS, 28 фактических CLI-задач той же EXE. Harness Python лишь запускает EXE и сверяет evidence. Интерактивное прохождение всех кнопок GUI: **NOT_RUN**, оставлено независимому Work и User E2E.

| Согласованный пункт | Проверка текущей EXE |
|---|---|
| 1. Запуск | frozen GUI smoke до/после |
| 2–4. Проект, паспорт, ВОМ, проверка | Гараж Казахстан; 9/2026, новый 10/2026, повтор 9/2026; каждый check PASS |
| 5. Приложение №1 | 104 видимые строки; 397 скрытых исключены; 288 шт.; 15822.14 кг; 1913.494 м²; M1 и порядок сверены после сохранения |
| 6. ВИК | Word DOCX создан и native read-only открыт |
| 7. УЗК | 33 единицы; SAMPLE_ONLY/неподтверждённый контроль сохранён |
| 8. Выбранные АКЗ | №0, пара документов; 0 не пустое значение |
| 9. Качество | Последнее; 2 страницы; пункты 12–14 по принятой логике; приложение АКЗ отдельно не перечислено |
| 10. Word/Excel | В каждом из 3 циклов 4 DOCX и 3 XLSX открыты native COM read-only без repair/save; SHA неизменны |
| 11. Manifest | Все 3 sealed manifests проверены; первый остаётся валиден после нового паспорта |
| 12. Official TEST STOP | 9/2026 запрещён; остатки неизменны |
| 13. Искусственное закрытие | 8001/2099 CLOSED_TEST, official_issue=false |
| 14. Однократное списание | 12/10 → 9/6; повтор ALREADY_CLOSED_TEST, остаётся 9/6 |
| 15. Восстановление | hard death во время публикации; RECOVERED_ABORTED, decrements=0, остатки 12/10; восстановлен точный исходный SHA |

Три последовательных комплекта — 21 документ. Native QUALITY page-1/page-2 первого комплекта визуально просмотрены. Остальные DOCX проверены COM-текстом/номером/датой/страницами, без заявления о просмотре всех страниц.

Первый тест readback после закрытия ошибочно сравнил raw OLE-container SHA и остановился на step 19. FAIL/traceback сохранены в e2e_2b9df014/FAILURE.json. `VBA_STREAM_READBACK.json` независимо прочитал все 11 OLE streams из registry/backup/staged: каждый поток, включая compiled VBA, совпал побайтно. Excel перепаковал только контейнер; accepted XLSM_PRESERVATION.json PASS. Исправлен тест Stage 9 на stream equality. Продолжено с первого незавершённого checkpoint, без повторной генерации трёх комплектов и без повторного первого списания.

Восстановленный XLSM SHA: `{e2e['restored_sha256']}`. Stage 8 backup/SHA, Windows lock, journal, idempotency и recovery проверены текущей EXE; защиты не отключались.

## 7. NEGATIVE TESTS

`EXE_NEGATIVE_FINAL.json`: {negative['pass']}/{negative['pass']} PASS, FAIL 0; ожидаемые FAIL приложения не скрыты и сохранены с traceback/result/supervisor.

| Сценарий | Test verdict |
|---|---|
{cases}

Дополнительно E2E проверяет repeat close и torn restore; lifecycle — normal/error/hard death; ownership-stop — неизвестный процесс. Все мутируемые файлы — собственные DEV/TEST fixtures.

## 8. SOURCE INTEGRITY

SOURCE_INTEGRITY_FINAL.json: {integrity['checked']}/{integrity['checked']} protected SHA без изменений. В baseline принятый Stage 1–8 code и реальные входные документы/реестры/справочники. Сохранена вся диагностическая 0.9. Новый EXE связывается с текущими raw/compiled исходниками через COMPILED_SOURCE_FINAL.json. Full aggregate hashes не заменяют индивидуальные списки.

## 9. CHANGED FILES

Ниже source/test files, отличающиеся от финального manifest 0.9; новые вспомогательные диагностические tests тоже перечислены. CHANGED_FILES.json содержит прежний/текущий SHA; STAGE9_DIFF_FINAL.patch — product diff.

{code_changes}

Дополнительно новые stage9_dist_091, stage9_build_091, STAGE9_DEV_091.md, эта validation-папка. DEV workspaces/diagnostics созданы вне production. Старые Stage 9 evidence не переписаны.

## 10. REMAINING FINDINGS

- Текущий воспроизводимый Office blocker устранён в проверенных сценариях. Точный OS-компонент прежнего AccessDenied остаётся UNKNOWN; необходимость administrator не доказана и не требуется новой сборкой.
- Независимый Work 09 NOT_RUN. Поддерживаемый launcher отдельного локального Work с доступом к Windows Office не предоставлен. create_thread local создаёт Codex; cloud Work требует явно выбранного cloud target и отдельно включённого local access, который здесь не подтверждён. Подмена независимого Work собственным Codex/субагентом не выполнена. [Условия локального доступа Work](https://learn.chatgpt.com/docs/enterprise/cloud-local-access).
- Полное GUI-click E2E, ручное открытие/закрытие viewer и пользовательская UX-приёмка Stage 9 NOT_RUN. Собственные native tasks не объявляются User E2E. Work должен проверить viewer повторно, включая shutdown и orphan snapshot.
- Реальные УЗК/АКЗ результаты, измерения/климат/партии/назначения для official issuance не подтверждены. Текущая EXE проверена с выбранным АКЗ 0; исторические synthetic проверки 1–3 не переносятся на этот frozen EXE.
- Принятые шаблоны привязаны к Гараж Казахстан. Другие проекты требуют отдельной принятой привязки. Application Python network guards активны; отдельная блокировка фоновой сети Office/физическое отключение сети NOT_RUN. ИИ/API для runtime не используются.
- Ранние FAIL сохранены: прежний non-child AccessDenied; R1 Word dialog; R2 ghost NativeOM binding; R3 cleanup; R4 invalid source (теперь AST preflight); R6 PowerShell missing in minimal PATH; R7 raw-container test assertion. Они не удалены и не названы PASS. PASS текущей SHA подтверждён отдельными FINAL evidence.

## 11–13. READINESS / STOP

READY FOR WORK 09 REVIEW: **YES** — текущая EXE и собственные обязательные проверки связаны SHA; готов WORK09_ASSIGNMENT.md.

READY FOR USER E2E: **NO** — необходим независимый Work PASS именно этой SHA, включая GUI viewer/UX.

READY FOR PRODUCTION: **NO** — DEV/TEST; independent/User acceptance и реальные official evidence/разрешение отсутствуют. Deployment, official closure и реальные списания не выполнялись.

EVIDENCE_MANIFEST.json записывается последним и связывает отчёт, задание, логи и referenced workspace/Office evidence. Самопроверки Codex не являются независимым review.

**STOP после передачи.**

## Приложение: фактические задачи E2E

| № | Task | Ожидаемый транспортный результат | Evidence |
|---|---|---|---|
{steps}
'''
(OUT/'STAGE9_REPORT.md').write_text(report,encoding='utf-8')
assignment=f'''# Work 09 — Финальная программа ВЗМК review

Независимый read-only review готовой Windows DEV/TEST программы. Не пересказывай отчёт Codex: самостоятельно исследуй продукт и воспроизведи проверки. Это задание для отдельного Work-исполнителя с локальным доступом к указанным файлам и установленному Office; Codex self-tests не считаются Work review.

## Точная привязка

ROOT: `{ROOT}`

EXE: `{EXE}`  
SHA256: `{build['exe_sha256']}`  
Onedir aggregate: `{build['aggregate_sha256']}` ({build['file_count']} files).  
Source/tests aggregate: `{source['aggregate_sha256']}`.

Прочитай STAGE9_REPORT.md, SOURCE_SHA256_FINAL.json, BUILD_SHA256_FINAL.json, COMPILED_SOURCE_FINAL.json, SOURCE_INTEGRITY_FINAL.json, CHANGED_FILES.json, STAGE9_DIFF_FINAL.patch и EVIDENCE_MANIFEST.json в `{OUT}`. Перед тестами заново проверь SHA EXE, всех bundled files и исходников. Не переноси PASS иной SHA.

Сохранённая диагностическая 0.9: SHA `{old['exe_sha256']}`; все {old['files']} files должны остаться неизменными. Принятые Stage 1–8 и {integrity['checked']} protected paths из PROTECTED_BEFORE.json неизменны.

## Границы

- READ-ONLY product/source/bundles/старое evidence. Не исправлять код и не пересобирать EXE. Findings вернуть Codex.
- Все записи только в новый отдельный DEV/TEST workspace/fixtures и новый review-evidence каталог. Не использовать текущие тестовые закрытия повторно и не менять их evidence.
- Не менять реальные реестры, паспорта, ВОМ, справочники, шаблоны, Stage 1–8. Не выполнять official closure/real UZK deduction/deployment, UAC или глобальный taskkill.
- Не использовать пользовательские Office. Если ownership неизвестен — STOP, не закрывать/не завершать неизвестный процесс. PID без creation time/image/исходного child handle недостаточен.
- Защиты TEST/SAMPLE_ONLY, locks, SHA, backup, journal, idempotency/recovery не отключать. Не придумывать реальные результаты/измерения.

## Самостоятельная проверка worker

1. Изучи old0.9 ownership.py:75–78 и новый ownership/supervisor/entrypoint. Проверь действительность original CreateProcess handles, PID/birth/image/launcher, STA references и native COM binding к точному child PID.
2. Сними собственный read-only baseline пользовательских Excel/Word с PID+creation time. Не используй чужие процессы для fault injection.
3. Текущая EXE: normal, handled error, hard death worker, повтор normal; проверь Workbook.Close/Document.Close(False), Quit, ожидание exit и orphan=0. Hard death должен остаться FAIL приложения с корректным test verdict. Normal/error должны обходиться без force.
4. Если force используется после hard death, проверь durable FORCE_BEGIN, исходный handle родителя, GetProcessId, birth/image и actual-child proof. Unknown ownership должен давать безопасный STOP.
5. Отдельно проверь GUI кнопку просмотра Word/Excel, два последовательных просмотра, закрытие viewer/главного окна и отсутствие orphan. Не считай native CLI verification доказательством всех GUI-событий.

## Полный frozen DEV E2E

Запускай саму указанную EXE без Python в PATH. Harness Python может управлять тестом, но не заменять runtime.

1. Запуск EXE/главного окна, создание нового DEV-стенда из ROOT; выбор Гараж Казахстан и 9/2026.
2. Выбор исходной ВОМ `{ROOT.parent/'Гараж казахстан/Паспорт/9 от 06.10.2026/ВОМ отгрузочный счет 183  (5).xls'}`.
3. Проверка: 104 visible, 397 hidden excluded, 288 qty, 15822.14 kg, 1913.494 m2, УЗК 33. Проверь все марки/порядок и connections M1 после XLSX-save.
4. Приложение №1, ВИК, SAMPLE_ONLY УЗК, выбранный АКЗ 0 (отличается от blank), качество последним. Никаких выдуманных factual control data.
5. Открой готовые 4 Word/3 Excel через native Office read-only без repair/save. Проверь номера/даты, TEST-label, итоги, 2-page quality и пункты 12–14; приложение АКЗ не перечисляется отдельно. Самостоятельно осмотри страницы.
6. Проверь manifest/SHA; новый паспорт 10/2026 на отдельной mutable DEV-копии, затем повтор 9/2026. Первый manifest должен остаться валиден. Несколько последовательных генераций.
7. Official close TEST/9/2026 запрещён без изменения registry.
8. На свежем искусственном fixture 8001/2099: preview → explicit TEST-confirm → CLOSED_TEST, official_issue=false, 12/10 → 9/6.
9. Повтор: ALREADY_CLOSED_TEST, 9/6 без второго списания. Проверь journal COMMITTED и exact backup.
10. На новом fixture hard death во время publish → recover/restore-preview → explicit TEST-confirm → RECOVERED_ABORTED, decrements=0, 12/10, exact исходный XLSM SHA, READY_TEST для нового осознанного запуска.
11. Перезапуск EXE, сохранность пользовательского Office/production, отсутствие своих orphan.

VBA: сравнивать побайтно **все OLE streams, включая compiled VBA**, имена/набор streams тоже. Не требовать неизменности только raw vbaProject.bin container, который Excel может перепаковать. Не ограничиваться textual macro source. Проверь accepted native.package — формулы/структура/стили сохраняются, только явно разрешённые save metadata/target cells исключены. Первоначальная ошибка теста и доказательство readback сохранены в e2e_2b9df014/FAILURE.json и VBA_STREAM_READBACK.json; самостоятельно проверь их вывод.

## Негативные проверки

Отсутствующая/невалидная ВОМ, tampered preview, SHA-конфликт собственного snapshot, неполный manifest, official flags, 9/2026, missing UZK, unconfirmed control, Windows exclusive lock, Excel before_save/save_io failure, damaged backup STOP. Дополнительно repeat close, torn publication restore, unknown ownership и worker hard death. Каждый FAIL приложения сопровождай независимым test verdict и проверь отсутствие опасной записи/orphan.

Приватные `--test-fault` hooks разрешены лишь artificial project + 8001/2099 + test_only=true; убедись, что их gate не обходит official protections. Исторические source-only АКЗ 1–3 не выдавай за frozen validation; при проверке новых synthetic variants явно обозначай fixtures.

## Воспроизводимость

Runtime не требует внешнего Python. Для автоматизированного harness допустим установленный тестовый Python 3.12 (только controller). Из ROOT задай STAGE9_VALIDATION_OUT в новый уникальный каталог, PYTHONUTF8=1, запускай `-B`. Скрипты сами ограничивают дочерний PATH до Windows. Не запускай legacy cleanup scripts или resume_e2e.py: они привязаны к прежним конкретным PID/checkpoint.

Последовательность: tests/stage9_091/lifecycle.py → e2e.py → negative.py → final_audit.py. Перед inventory.py скопируй неизменяемые PROTECTED_BEFORE.json и OLD_09_BUILD_BEFORE.json в новый review output. final_audit/сохранение FINAL используют exclusive-create; для повторного независимого прогона нужен новый output. GUI tests в tests/stage9_091/gui.py имеют фиксированный output Codex; не перезаписывай его — создай собственный harness/новое evidence либо проверь UI вручную.

Не запускать build.py, make_handoff.py, seal_evidence.py из review: продукт/readiness и canonical handoff не менять. Сетевые Python guards активны; Office background network isolation отдельно не доказана. Отделяй подтверждённое от NOT_RUN.

## Отчёт Work / STOP

Верни exact reviewed EXE/source/onedir SHA; собственные команды и evidence; BLOCKER/MAJOR/MINOR с путями/строками/воспроизведением; worker/GUI/Office/frozen full E2E/negative/Stage8 fidelity/source integrity verdicts; все FAIL/NOT_RUN. Обязательно реши, устранён ли прежний blocker без admin и закрытия чужого Office.

PASS допускается только по собственным проверкам конкретной SHA. Итог READY FOR USER E2E — YES/NO с обоснованием; READY FOR PRODUCTION остаётся NO (DEV/TEST, official controls и отдельного разрешения нет). Если найден defect, не исправляй его. Передай findings для разрешённого Stage9-only исправления и нового SHA/targeted review.

После отчёта STOP. Независимый review сейчас **NOT_RUN**; этот файл является заданием, не результатом Work.
'''
(OUT/'WORK09_ASSIGNMENT.md').write_text(assignment,encoding='utf-8')
print('HANDOFF WRITTEN',build['exe_sha256'],source['aggregate_sha256'],len(changed),'changed source/test files')
