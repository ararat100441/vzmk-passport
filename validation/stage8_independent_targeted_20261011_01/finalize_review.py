"""Seal the bounded independent review; does not alter product or prior evidence."""
import hashlib,json
from pathlib import Path
RUN=Path(__file__).resolve().parent;ROOT=RUN.parents[1]
def read(p):return json.loads(Path(p).read_text(encoding='utf-8-sig'))
def save(p,v):Path(p).write_text(json.dumps(v,ensure_ascii=False,indent=2),encoding='utf8')
def sha(p):return hashlib.sha256(Path(p).read_bytes()).hexdigest()
def link(label,path,line=None):return '['+label+'](<'+str(path)+((':'+str(line)) if line else '')+'>)'
base=read(RUN/'RESULTS.json');cases={x['name']:x|{'evidence':str(RUN/'cases'/(x['name']+'.json'))} for x in base['results']}
substitutions=[]
for batch in ('rerun_03','rerun_04'):
    for x in read(RUN/batch/'RESULTS.json')['results']:
        substitutions.append({'case':x['name'],'superseded':cases[x['name']]['evidence'],'replacement':str(RUN/batch/'cases'/(x['name']+'.json'))})
        cases[x['name']]=x|{'evidence':str(RUN/batch/'cases'/(x['name']+'.json'))}
probe=read(RUN/'lifecycle_probe_90/RESULT.json');cases[probe['name']]=probe|{'evidence':str(RUN/'lifecycle_probe_90/RESULT.json')}
passed=sum(x['status']=='PASS' for x in cases.values());failed=sum(x['status']=='FAIL' for x in cases.values())
assert len(cases)==35 and passed==34 and failed==1
assert [x['name'] for x in cases.values() if x['status']=='FAIL']==['partial_XLSM_supported_operator_recovery']
integrity=read(RUN/'SOURCE_INTEGRITY.json');assert integrity['status']=='PASS' and integrity['protected_files']==22627 and integrity['current_product_matches']
initial=read(RUN/'INITIAL_BINDING.json');assert initial['delivery_matches']
matrix={'PASS':passed,'FAIL':failed,'total':len(cases),'results':list(cases.values()),'superseded_checks':substitutions,
        'history_preserved':['harness_01','RESULTS.json','rerun_03/RESULTS.json','lifecycle_probe/RESULT.json'],
        'reasons':['harness_01: fixture JSON newline changes required rebinding references; product rejected inputs before INTENT',
                   'native publisher: nonterminated ctypes name created a sibling path; corrected buffer + exact destination SHA readback',
                   '25-second process-exit assertion was too short; native isolated client proved normal delayed exit at 60.03 seconds; dependent cases rerun with 90-second bound'],
        'product_unchanged':True,'verdict':'FAIL','ready_for_user_e2e':'NO','findings':{'BLOCKER':0,'MAJOR':1,'MINOR':0}}
save(RUN/'FINAL_TARGETED_RESULTS.json',matrix)
product=read(RUN/'PRODUCT_AFTER.json')
changes='\n'.join('| '+Path(x['path']).name+' | `'+x['after']+'` |' for x in initial['product_changes'])
rows='\n'.join('| '+x['name']+' | '+x['status']+' |' for x in cases.values())
doc=f'''# ВЗМК — Этап 8 — Final Targeted Independent Review

11.10.2026, Europe/Samara. Только изолированный DEV/TEST, без исправления product.

**FINAL VERDICT: FAIL. READY FOR USER E2E: NO. BLOCKER 0 / MAJOR 1 / MINOR 0.**
Четыре прежних FAILED-сценария теперь PASS. Новый обязательный критерий восстановления
частичного XLSM не выполнен: поставка обнаруживает аварию и сохраняет backup, но не
содержит документированного операторского пути возврата реестра в работу.
Итоговая независимая матрица: **{passed} PASS / {failed} FAIL / {len(cases)} проверок**.

## 1. ACTUAL PRODUCT STATE

Проверена текущая версия в `{ROOT}`. Движок по-прежнему допускает только
искусственный `CLOSED_TEST`, `official_issue=false`; официального адаптера нет.
Прочитаны оба исходных FINAL_REPORT.md, текущие изменённые модули, неизменённые
model/fixtures/CLI, diff и привязанные протоколы. Полный аудит Stage 1–7 не повторялся.
Semantic index вернул INTERNAL_ERROR; применено прямое чтение точных исходников.

Fix manifest SHA-256: `aa073f6771baf0c68d18a8a6f71c4e4030a742c5627f3a1592975ea47dcff5c4`.
Независимо совпали **12/12 product**, **9/9 runtime**, **7/7 новых тестов**,
**2984/2984 evidence files**. Fingerprint поставки:
`599342381eac6cce617c2d474608be16dcf17034387b9480d69b2bc800f3c8a3`.
Канонические 38 PASS действительно связаны с указанными snapshots; они не использованы
как замена независимому испытанию. Доказательства: INITIAL_BINDING.json,
PRIOR_EVIDENCE_BINDING.json, PRODUCT_BEFORE.json / PRODUCT_AFTER.json.

## 2. EXACT CODE CHANGES

По самостоятельному diff от before_product изменены ровно четыре модуля и инструкция:
core.py связывает preview с file identity и держит guard до COMMITTED/recovery;
storage.py убрал share-delete/rename publication и пишет через удерживаемый handle,
защищая target и ancestors; office.py разделил cleanup и добавил ownership/START/END;
native.py сохраняет primary Save HRESULT, заранее связывает Save и диагностирует Close.
STAGE8_DEV.md теперь честно описывает неатомарную запись и STOP. Fixtures/model/gui,
Stage 1–7 и launcher не менялись относительно fix checkpoint.

| Изменённый файл | Текущий SHA-256 |
|---|---|
{changes}

Exact diff: INDEPENDENT_PRODUCT.diff. Во время этой проверки product не изменялся;
полные SHA всех 12 файлов совпали до/после и проверялись после каждого сценария.

## 3. CONCURRENT REPLACEMENT TESTS

PASS в проверенном Windows/local NTFS scope. Native Microsoft Excel заранее сохранил
корректную конкурентную книгу со значением **4**; normal reopen подтвердил её.
Отдельный процесс, игнорирующий application lock, пытался выполнить rename+replace,
native FileRenameInfoEx/POSIX replace и прямую запись. Проверены before_replace,
между блоками копирования и after_replace до COMMITTED; переименование registry parent
также заблокировано. Windows отверг конкурентные операции; Stage 8 закончил **12→9**,
вторая марка **10→6**, identity прежняя, COMMITTED ровно один.

Если внешняя native-публикация успела **до acquisition**, действительно осталось **4**;
close старого preview дал STOP, journal пуст. Byte-identical замена с новой file identity
тоже дала STOP. Два одновременных worker дали одно закрытие/списание. Уже открытые writer
и writable mapping запретили acquisition. Дополнительное TOCTOU release/reacquire окно
на участке публикации по исходникам отсутствует: один handle живёт до выхода из _commit.

Доказательства: EXTERNAL_NATIVE_4.json; FINAL_TARGETED_RESULTS.json, concurrency cases
в cases/ и rerun_03/cases/. API publisher проверяет SHA именно целевого пути при успехе.

## 4. PREVIOUS FAILED TESTS — NOW

| Прежний FAILED-сценарий | Теперь | Результат |
|---|---|---|
| external_rename_then_replace_before_publish | PASS | WinError 32; 12→9; один COMMITTED |
| external_native_atomic_replace_before_publish | PASS | native replace заблокирован |
| concurrent_valid_xlsm_balance_resurrection | PASS | корректная внешняя книга 4 не опубликована; штатно 9 |
| cleanup_com_failure_must_still_attempt_quit | PASS | первая ошибка записана; Close и Quit вызваны |

Прежние конкретные MAJOR и MINOR устранены в испытанном объёме. Новое замечание M2
относится к аварийной восстановимости изменённого протокола публикации.

## 5. XLSM CRASH CONSISTENCY

Реальный worker завершался `os._exit(73)` до Save, после первого блока publication
и после полной публикации. Fault hook только задаёт точку смерти; запись XLSM и
native Save настоящие. Произвольная смерть во время copy моделируется фактической
частичной записью через product handle, а не подменой результата recovery.

| Вопрос | Независимый результат |
|---|---|
| Обнаружение partial XLSM | SHA отличается от before/after; independent ZIP разбор выявил CRC failure |
| Ошибочный successful save/COMMITTED | Нет: STAGED относится к отдельной целой копии; target остаётся APPLYING |
| Повторное списание | preview, close и fresh-process recovery блокируются; журнал/книга не меняются |
| Исходный backup | SHA и metadata совпадают; native normal read-only open, исходные 12/10 и 11 VBA streams |
| Новый запуск | AMBIGUOUS_OPERATION_STOP_NO_REPLAY; дальнейшие операции запрещены |
| Открытие повреждённой книги приложением | native.inspect отверг package до DispatchEx: native_opens=0 |
| Смерть до Save | неизменённая книга; fresh recovery ABORTED_UNCHANGED |
| Смерть после полной публикации | recovery дописал COMMITTED, registry_writes=0, остаток 9 |
| Вернуть partial книгу в работу штатным способом | FAIL, см. M2 и раздел 6 |

Неатомарность подтверждена, но не сопровождалась ложным закрытием или потерей backup
в этих сценариях. Это не аттестация аппаратного отключения питания.

## 6. BACKUP RECOVERY PROCEDURE

**Доступная поставка не содержит проверяемого документированного операторского workflow
для partial XLSM.** `recover` меняет только журнал: прежний SHA → ABORTED, точный after
при APPLYING → COMMITTED, иначе STOP. Автоматическое восстановление само по себе не
обязательно; отсутствует и документированный ручной порядок безопасного возврата.

На аварийном стенде отдельно проверена обычная замена файла копией проверенного backup.
Исходные байты и остатки **12/10** вернулись, однако file identity изменилась;
fresh recovery выдал `AMBIGUOUS_REGISTRY_FILE_ID_STOP_NO_REPLAY`, операция осталась pending.
Удалять/перепечатывать journal или SANDBOX для обхода запрета нельзя считать восстановлением.

На **другом** новом crash-fixture review-only код удержал sandbox lock и RegistryGuard,
повторно проверил backup/metadata, сохранил повреждённые байты для evidence и скопировал
backup через тот же file handle с проверками SHA/identity. Fresh recovery затем дал
ABORTED_UNCHANGED; новый явный preview/close списал один раз до 9. Это доказывает
восстановимость сохранённых данных и техническую возможность решения. Данный внутренний
эксперимент не является доставленной операторской командой/инструкцией и не закрывает M2.
Product и его эксплуатационная инструкция в review не исправлялись.

Доказательства: RECOVERY_GAP.json, operator_replace_backup_attempt.json,
expert_fixture_restore_proves_backup_data_recoverable.json, DAMAGED_XLSM_EVIDENCE.bin
в соответствующих fixtures. Требуется поставляемый, документированный и независимо
проверенный порядок восстановления с backup binding, сохранением evidence, защитой
от конкуренции, согласованием identity/journal и новым явным подтверждением при abort.

## 7. COMMITTED / IDEMPOTENCY SAFETY

PASS для проверенных ветвей. COMMITTED следует после полной copy/truncate/flush,
SHA и identity check под guard. При partial запись остановилась на APPLYING, false
COMMITTED отсутствовал. Повтор старого preview, иной путь запроса/отгрузка того же
project/passport и fresh worker вернули ALREADY_CLOSED_TEST без изменения SHA/журнала.
Два worker не списали дважды. Успешный recovery после полной публикации не писал XLSM.
No-UZK закрытие создало только journal, без backup/stage/записи реестра.

## 8. EXCEL COM CLEANUP

PASS в окончательных проверках. На реальном собственном Excel независимо внедрены
ошибки CutCopyMode, Workbooks.Close и Quit. Все три попытки выполнялись; ERROR сохранялся
в END.json; исходный S.Stop и исходный Save HRESULT не подменялись. Native Workbook.Close
ошибка записывалась отдельно, с note на primary exception. Для намеренного отказа Quit
review отдельно вызвал настоящий Quit только у захваченного собственного instance.

Unrelated TEST Excel с sentinel ячейкой пережил работу другого instance, значение не
изменилось. Пользовательские Excel не закрывались; исходный список Excel был пуст.
При os._exit finally не требуется: собственный crash orphan обрабатывался только по
START PID+creation time и ROT-пути книги; при необходимости — адресно, без global taskkill.
Это review-only cleanup, не способность product после собственной смерти.

Выход Excel бывает задержан: чистый native.inspect без mock завершил собственный
процесс через **60.03 sec**, пока клиент Python продолжал работать. Первоначальный
25-second timeout был недостаточен; три зависевших от него проверки повторены с
90-second bound. Финальный ownership/process snapshot: собственных live Excel **0**.
Искусственно короткий timeout не оформлен как новый MINOR.

## 9. HRESULT 0x80010108

`0x80010108` — RPC_E_DISCONNECTED, разрыв COM-объекта с клиентами.
[Microsoft COM error codes](https://learn.microsoft.com/en-us/windows/win32/com/com-error-codes-3).
Точный HRESULT внедрён в Save на настоящем собственном Excel; отдельно реальный
Excel.Quit разорвал COM перед связанным Save. Оба случая: STOP, исходный XLSM прежний,
INTENT/BACKUP_VERIFIED, backup verified, COMMITTED нет, fresh recovery ABORTED_UNCHANGED,
повторного списания нет. Primary HRESULT/cause и cleanup diagnostics сохранены.
Причина первоначального спонтанного события **UNKNOWN**; её исправление не заявляется.

## 10. XLSM PRESERVATION

PASS. Native Microsoft Excel 16.0 normal open, CorruptLoad=0, формат 52, read-only,
HasVBProject, отключённые VBA/events/link updates. Независимый OLE-разбор подтвердил
побайтную идентичность всех **11 VBA streams** у backup и результата. Native значения:
12/10 → 9/6; S8_FORMULAS veryHidden, SUM 30→23, =2+3 остаётся 5; выражения формул прежние.

Независимый raw OOXML разбор: 13 parts, прежние имена частей; styles и все relationships
побайтно одинаковы; sheets/definedNames/protection/externalReferences идентичны.
Ровно три изменившиеся ячейки: C24, C25 и зависимый formula cache A1. Остальные различия
parts — допустимые save metadata и VBA container при неизменных streams.
Backup после настоящей смерти worker также нормально открыт native Excel с исходными
остатками и VBA. Рабочая УЗК.xlsm не использовалась для испытаний.
Доказательства: NATIVE_PRESERVATION.json, INDEPENDENT_RAW_XLSM.json, crash cases.

## 11. TARGETED TEST RESULTS

**{passed} PASS / {failed} FAIL / {len(cases)} итоговых acceptance checks.**
Четыре ранее проваленных теста — 4/4 PASS. Ограниченная регрессия включает one-time debit,
reclose, no-UZK, insufficient, changed SHA, corrupt verified backup, corrupt journal,
safe abort/finalize, partial STOP и запрет 9/2026.

| Case | Result |
|---|---|
{rows}

FINAL_TARGETED_RESULTS.json указывает точный source протокол каждого итогового случая.
История не удалена: harness_01 (8 PASS / 26 FAIL, ошибка SHA rebinding искусственных JSON),
первый native batch RESULTS.json (26 PASS / 8 FAIL), rerun_03 (4 PASS / 3 FAIL),
rerun_04 (3 PASS / 0 FAIL), короткий lifecycle probe и окончательный lifecycle_probe_90.
Native publisher первого batch не завершал WCHAR буфер NUL: создавался sibling с лишними
символами, а не конкурирующая версия целевого пути. Исправлен только harness: NUL buffer
и обязательный SHA readback именно destination; четыре зависевших случая повторены.
Три другие временные ошибки были слишком коротким ожиданием выхода Excel. Промежуточные
результаты не выдаются за product дефекты или окончательные PASS; superseded_checks
явно связывает замены. Runtime product во всех запусках битово идентичен.

## 12. SOURCE INTEGRITY

**До/после текущей проверки: PASS, {integrity['protected_files']} уникальных protected files,
0 изменённых/нечитаемых; все 12 текущих product files прежние.** Scope объединяет весь
доступный workspace, прежние receipts/backup/validation и весь заявленный список 19618,
включая доступные внешние рабочие источники. Новая review directory исключена; pycache/.git
не аттестованы. Paths/SHA: PROTECTED_DATA_BEFORE.json, PROTECTED_DATA_AFTER.json,
SOURCE_INTEGRITY.json. В исходном расширенном snapshot 22628 путей изменился только
один служебный `.codebase-index/...lease/heartbeat...json` фонового worker. Его прежний
и новый SHA сохранены; он явно исключён как volatile tool metadata. Нефильтрованная
сверка сохранена в ALL_ACCESSIBLE_INTEGRITY.json и PROTECTED_BEFORE/AFTER.json; это
изменение не объявляется неизменным product/source. Остальные 22627 совпали точно.

Исторические fix PROTECTED_BEFORE и final_integrity_02/PROTECTED_AFTER действительно
равны для **19618/19618**. При текущей независимой сверке с историческим checkpoint
совпали **19584**, а **34** уже отличались **до** тестов. Все 34 находятся исключительно
в `.cache/codex-runtimes/...` (bundled runtime/cache); причина их изменения отдельно
не устанавливалась. Поэтому утверждение «все 19618 до сих пор совпадают со старым SHA»
не повторяется. Регистры, ВОМ, документы, справочники, Stage 1–7 и старые backup/evidence
из списка совпали с историческими SHA. Текущие SHA этих 34 cache files тоже включены
в новую проверку до/после. Доказательства: CLAIMED_19618_CURRENT.json и INITIAL_BINDING.json.

Before manifest SHA: `{integrity['before_manifest_sha256']}`.
After manifest SHA: `{integrity['after_manifest_sha256']}`.

## 13. FINDINGS — BLOCKER 0 / MAJOR 1 / MINOR 0

**M2 — MAJOR: partial publication оставляет рабочий XLSM непригодным без поставленного
операторского восстановления; обычный restore backup упирается в identity STOP.**

Файлы/строки: {link('storage.py:271',ROOT/'stage8_closure/storage.py',271)} — запись
сразу в target до flush/проверки; {link('core.py:172',ROOT/'stage8_closure/core.py',172)} —
STOP при смене identity; {link('core.py:188',ROOT/'stage8_closure/core.py',188)} — partial SHA
имеет только STOP; {link('STAGE8_DEV.md:71',ROOT/'STAGE8_DEV.md',71)} — инструкция содержит
только journal finalize/abort/STOP, без безопасного workflow восстановления файла.

Причина: новый in-place copy может разрушить оригинал до COMMITTED, тогда как recovery
проектируется как journal-only. Backup целый, но его замена создаёт другую identity,
которую recovery отвергает; поддержанного manual restore/reconciliation нет.
Следствие: после воспроизведённого обычного process crash реестр и дальнейшие операции
остаются недоступны оператору, требуется недокументированное вмешательство через
внутренние API. STOP защищает от double-debit, однако не восстанавливает рабочий учёт.

Это **не BLOCKER утраты данных**: backup доказан, false COMMITTED нет, expert fixture
показал техническую восстановимость. Но обязательный recovery критерий и готовность
передать продукт оператору не выполнены. Автоматический rollback не навязывается;
нужен проверяемый поставляемый порядок безопасного восстановления и targeted rereview.

## 14. REMAINING LIMITATIONS

In-place publication не атомарна: внешний reader может увидеть partial bytes. Product
проверяет собственный state/package, но сторонний Excel-reader не координируется с journal.
Crash orphan требует ownership review, автоматического PID termination нет. Выход Office
может задерживаться примерно на минуту. Локальная SHA chain не защищает от намеренной
перепечатки всех evidence администратором. Идемпотентность ограничена одним sandbox/registry.
Старые pending plans без identity требуют отдельного review. Исходная spontaneous COM
причина UNKNOWN. M2 остаётся открытым.

## 15. NOT_RUN

Официальное закрытие/реальные списания/выдача, ручной пользовательский E2E, повторная
полная GUI-проверка, полная регрессия Stage 1–7, аппаратное отключение питания, другие ФС/SMB,
production, Stage 9, EXE, PDF/deployment — NOT_RUN. Настоящая смерть Python именно внутри
native Save в этой targeted rereview повторно не выполнялась: проверены реальные смерти
до Save/во время publication/после publication и отдельные COM Save отказы; прежний Save
callback evidence сохранён и SHA связан. Это не расширенный полный аудит.

## 16. FINAL VERDICT: FAIL

Конкурентная запись и прежнее скрытие cleanup-ошибки исправлены в испытанном объёме.
Безопасное обнаружение partial write и сохранность backup доказаны. Обязательный
операторский recovery после partial write не поставлен; M2 не позволяет принять Этап 8.
Исправления в рамках этой проверки не выполнялись.

## 17. READY FOR USER E2E: NO

До поставки и независимой проверки восстановления partial XLSM готовность не подтверждена.
Реальных списаний 0. **STOP. Stage 9 не начат. EXE/deployment не выполнялись.**
'''
(RUN/'FINAL_REPORT.md').write_text(doc,encoding='utf8')
manifest={'schema':'STAGE8_TARGETED_INDEPENDENT_REVIEW_V1','verdict':'FAIL','ready_for_user_e2e':'NO','product':product,
          'review_evidence':{str(p.resolve()):sha(p) for p in sorted(RUN.rglob('*')) if p.is_file() and p.name not in ('REVIEW_MANIFEST.json','SEAL_VERIFICATION.json') and '__pycache__' not in p.parts}}
save(RUN/'REVIEW_MANIFEST.json',manifest)
errors={p:'SHA_MISMATCH' for p,h in (manifest['product']|manifest['review_evidence']).items() if sha(p)!=h}
save(RUN/'SEAL_VERIFICATION.json',{'status':'PASS' if not errors else 'FAIL','product':len(product),'evidence':len(manifest['review_evidence']),'errors':errors,'manifest_sha256':sha(RUN/'REVIEW_MANIFEST.json')})
assert not errors
print('FINAL',passed,failed,'MANIFEST',sha(RUN/'REVIEW_MANIFEST.json'),'EVIDENCE',len(manifest['review_evidence']))
