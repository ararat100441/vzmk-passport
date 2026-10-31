"""Seal the new bounded evidence package last, after successful final tests."""
import gc,json,sys
from pathlib import Path
import win32api,win32con,win32process,win32event
from stage8_closure import core,storage as S
RUN=S.ROOT/'validation/stage8_fix_20261010_01'
base=S.read(RUN/'targeted_03/RESULTS.json')
supplement=S.read(RUN/'targeted_04/RESULTS.json')
delivery_test=S.read(RUN/'targeted_05/RESULTS.json')
assert base['FAIL']==1 and base['PASS']==35
failed=[r for r in base['results'] if r['status']=='FAIL']
assert failed[0]['case']=='real_RPC_E_DISCONNECTED_Save'
assert supplement['FAIL']==0 and supplement['PASS']==2
assert delivery_test['FAIL']==0 and delivery_test['PASS']==1
final_case=delivery_test['results'][0]|{'case':'final_delivery_native_normal_reopen','source_run':'targeted_05'}
matrix={'PASS':38,'FAIL':0,'results':[r for r in base['results'] if r['status']=='PASS']+supplement['results']+[final_case],
    'sources':['targeted_03/RESULTS.json','targeted_04/RESULTS.json','targeted_05/RESULTS.json'],
    'superseded_failed_check':{'path':'targeted_03/cases/real_RPC_E_DISCONNECTED_Save.json',
        'reason':'Exact HRESULT assumption replaced by real-server-death + explicitly injected 0x80010108 checks; historical FAIL retained.'}}
S.write(RUN/'CANONICAL_TARGETED_RESULTS.json',matrix)
before=S.read(RUN/'targeted_03/PRODUCT_BEFORE.json');after=S.read(RUN/'targeted_03/PRODUCT_AFTER.json')
assert before==after==S.read(RUN/'targeted_04/PRODUCT_BEFORE.json')==S.read(RUN/'targeted_04/PRODUCT_AFTER.json')==S.read(RUN/'targeted_05/PRODUCT_BEFORE.json')==S.read(RUN/'targeted_05/PRODUCT_AFTER.json')==core.code_bindings()
for name in ('targeted','worker'):
    (RUN/'targeted_05'/(name.upper()+'_SNAPSHOT.py')).write_bytes((S.ROOT/'tests/stage8/work08_fix'/(name+'.py')).read_bytes())
integrity=S.read(RUN/'final_integrity_02/SOURCE_INTEGRITY.json');assert integrity['status']=='PASS'
changes=S.read(RUN/'final_integrity_02/CHANGED_PRODUCT_SHA.json');assert len(changes)==5
checkpoint=S.read(RUN/'PRODUCT_CHECKPOINT.json')
delivery_before=S.read(RUN/'targeted_05/DELIVERY_BEFORE.json')
delivery_after=S.read(RUN/'targeted_05/DELIVERY_AFTER.json')
delivery_current={p:S.sha(p) for p in checkpoint}
assert delivery_before==delivery_after==delivery_current
(RUN/'after_product').mkdir(exist_ok=False)
for p,h in delivery_current.items():
    snapshot=RUN/'after_product'/Path(p).name
    snapshot.write_bytes(Path(p).read_bytes())
    assert S.sha(snapshot)==h
S.write(RUN/'FULL_PRODUCT_BINDING.json',{'before_tests':delivery_before,'after_tests':delivery_after,
    'current':delivery_current,'fingerprint_sha256':S.digest(delivery_current)})

# No session, workbook or native writer is opened during sealing.
live=[];owned=[]
for p in RUN.rglob('START.json'):
    record=S.read(p)
    if record.get('schema')!='STAGE8_OWNED_EXCEL_V1' or not record.get('owned'):continue
    owned.append({'path':str(p),'pid':record['pid'],'creation_time':record['creation_time']})
    try:h=win32api.OpenProcess(win32con.PROCESS_QUERY_INFORMATION|win32con.SYNCHRONIZE,False,record['pid'])
    except Exception:continue
    try:
        if (win32process.GetExitCodeProcess(h)==259 and
                win32process.GetProcessTimes(h)['CreationTime'].isoformat()==record['creation_time']):
            win32event.WaitForSingleObject(h,25000)
            if win32process.GetExitCodeProcess(h)==259:
                live.append({'path':str(p),**record})
    finally:h.Close()
S.write(RUN/'FINAL_EXCEL_OWNERSHIP.json',{'status':'PASS' if not live else 'FAIL',
                                      'owned_sessions':len(owned),'live_owned':live})
assert not live,live

files=[]
for root in sorted((RUN/'sandboxes').iterdir()):
    meta=root/'SANDBOX.json'
    if not meta.exists():continue
    registry=root/'registry/УЗК_TEST.xlsm';h=S.history(root)
    files.append({'sandbox':str(root),'initial_registry_sha256':S.read(meta)['initial_registry_sha256'],
        'final_registry_sha256':S.sha(registry),'final_registry_identity':S.file_identity(registry),
        'inputs':{str(p):S.sha(p) for p in sorted((root/'inputs').glob('*.json'))},
        'journal':{str(p):S.sha(p) for p in sorted((root/'journal').glob('*.json'))},
        'states':[e['state'] for e in h['events']], 'closed_keys':list(h['closed']),
        'pending':h['active']['state'] if h['active'] else None,
        'backups':{str(p):S.sha(p) for p in sorted((root/'backups').rglob('*')) if p.is_file()}})
S.write(RUN/'TEST_INPUT_OUTPUT_JOURNAL_SHA.json',files)
S.write(RUN/'PRODUCT_CURRENT.json',core.code_bindings())
newtests={str(p):S.sha(p) for p in sorted((S.ROOT/'tests/stage8/work08_fix').glob('*.py'))}
S.write(RUN/'NEW_TEST_FILES_SHA.json',newtests)

sha_table='\n'.join('| '+Path(c['path']).name+' | `'+c['before_sha256']+'` | `'+c['after_sha256']+'` |' for c in changes)
test_table='\n'.join('| '+r['case']+' | '+r['status']+' |' for r in matrix['results'])
new_table='\n'.join('| '+str(Path(p).relative_to(S.ROOT))+' | `'+h+'` |' for p,h in newtests.items())
report=f'''# ВЗМК — Этап 8: targeted fix after independent Work 08

Дата: 2026-10-10, Europe/Samara. Машинные timestamps в JSON — UTC.
Только изолированный DEV/TEST. Полный аудит проекта не повторялся.
Рабочая папка: `{S.ROOT}`.
Новый пакет evidence: `{RUN}`.

## 1. ROOT CAUSE MAJOR

Подтверждено до изменения product на трёх новых стендах: открытый read handle
разрешал FILE_SHARE_DELETE, поэтому защищал прежний file object от записи,
но разрешал внешнему процессу сменить целевой путь. Последний SHA предшествовал
FileRenameInfoEx/POSIX rename. Чужая версия со значением 4 исчезала под старым
расчётом 12−3=9, после чего возникал COMMITTED.
Точные исходные product snapshots: before_product/; исходные SHA:
PRODUCT_CHECKPOINT.json. Три воспроизведения: before_reproduction/*.json.

## 2. ROOT CAUSE MINOR

CutCopyMode, Workbooks.Close и Quit находились внутри общего try/except pass.
Первая ошибка пропускала остальные действия и исчезала. На собственном реальном
Excel до исправления воспроизведён живой PID и отсутствующий Quit.
Доказательство: before_reproduction/cleanup.json.

## 3. EXACT CODE CHANGES

Изменены четыре существующих файла кода Stage 8: storage.py, core.py,
office.py, native.py. Дополнительно необходимо исправлена инструкция STAGE8_DEV.md:
её строки 39–41 и 81–82 продолжали обещать прежнее атомарное переименование.
В инструкции теперь явно указаны handle-copy, partial-write STOP и ограничения cleanup.
fixtures/model/gui, Stage 1–7, launcher и шаблоны не менялись.
Exact unified diff: final_integrity_02/EXACT_PRODUCT.diff.
Точные файлы окончательной поставки сохранены в after_product/.

storage.py: frozen больше не разрешает DELETE; добавлены file identity,
RegistryGuard и удержание directory ancestors; старый POSIX replacement удалён.
core.py: preview связывает volume serial/file index; close и recover удерживают
guard; публикация выполняется через тот же handle; identity сверяется при recovery
и повторном обращении к закрытому журналу. Незавершённый legacy plan без identity
останавливается для review, старые receipts не переписываются.
office.py: независимые cleanup attempts, START/END diagnostics, проверка владения
через новый PID/creation time, сохранение primary exception и освобождение локальных
application references перед cleanup STOP.
native.py: диагностика Workbook.Close, сохранение Save exception при ошибке Close,
заранее связанный Save для сохранения HRESULT при disconnect; native writers только
на частной staged копии.

## 4. CONCURRENCY PROTECTION

Только локальный fixed NTFS; иные файловые системы отклоняются. Registry handle
открывается READ|WRITE с FILE_SHARE_READ, без WRITE/DELETE sharing. Он не закрывается
и не переоткрывается между последней проверкой, записью и COMMITTED. Все directory
ancestors также удерживаются без WRITE/DELETE sharing; reparse/hardlink запрещены.
Сверяются volume serial + file index и SHA содержимого через удерживаемый handle.
Это OS-enforced защита, независимая от application lock. Уже существующие writer
и writable mapping отклоняют acquisition; будущие write/rename/replace блокируются.

Гарантии API проверены по официальным
[CreateFileW](https://learn.microsoft.com/en-us/windows/win32/api/fileapi/nf-fileapi-createfilew),
[GetFileInformationByHandle](https://learn.microsoft.com/en-us/windows/win32/api/fileapi/nf-fileapi-getfileinformationbyhandle),
[FILE_RENAME_INFORMATION](https://learn.microsoft.com/en-us/windows-hardware/drivers/ddi/ntifs/ns-ntifs-_file_rename_information)
и фактически испытаны отдельным процессом на текущем Windows/NTFS:
before_reproduction/NO_DELETE_HANDLE_PROBE.json; targeted_03/cases/*.
Среда: WINDOWS_ENVIRONMENT.json. Нет заявлений о гарантиях других ФС/SMB.

## 5. RACE REPRODUCTION BEFORE/AFTER

BEFORE: 12 → внешняя native-saved версия 4 → Stage 8 пишет 9, COMMITTED:
три подтверждённых прежних дефекта (rename/native/numeric).
AFTER: valid native XLSM со значением 4 подготовлен и отдельно открыт Excel;
внешняя публикация в before_replace блокируется ОС, Stage 8 завершает 12−3=9.
Тест work08_native_before_publish одновременно проверяет точный numeric 12/4/9.
rename+replace и обычный os.replace проверены отдельно. Если native replace успевает
до acquisition, остаётся 4; close выдаёт STOP без INTENT/COMMITTED.
Дополнительно write/native replace/parent rename проверены после INTENT, прямо
между блоками копирования и после записи перед COMMITTED. Ни один publisher
не обходит app lock добровольно: это отдельный процесс с собственным Win32 API.

## 6. JOURNAL / COMMITTED SAFETY

APPLYING записывается до публикации. COMMITTED следует только после полного
copy/truncate/FlushFileBuffers и проверки SHA + прежней file identity под guard.
Два одновременных процесса: одно закрытие/одно списание; повторный вызов не пишет
XLSM и не добавляет COMMITTED. No-UZK путь не меняет registry и не создаёт backup.
Неоднозначный partial copy сохраняет APPLYING; preview/close/recover дают STOP.

## 7. EXCEL COM CLEANUP

Каждый CutCopyMode/Workbooks.Close/Quit имеет собственный try, ошибки остаются
в END.json. Ошибка Workbook.Close также записывается; primary Save error остаётся
cause, включая HRESULT и notes. Cleanup failure на обычном пути даёт диагностический
STOP; при primary exception не подменяет её. Неподтверждённый PID не закрывается.
START.json содержит PID, creation time, expected workbook paths до открытия книги.
Нормальное завершение и controlled exceptions испытаны на реальном Excel.
Отдельный unrelated TEST Excel пережил работу Stage 8 с неизменённой ячейкой.
FINAL_EXCEL_OWNERSHIP.json: собственных живых экземпляров после тестов 0.

Принудительный os._exit обходит finally. Orphan действительно воспроизведён при
смерти внутри Save и до Save, а также возможен после уже вызванного Quit при живых
COM references. Для TEST выполнена контролируемая обработка: только PID с точным
creation time из своего START; ROT binding без GetObject activation; проверены
пути всех книг и invisible instance, Close без Save и Quit с grace period до 25 секунд.
Если процесс остался,
использован только подтверждённый собственный PID/handle; глобального taskkill нет.
Для post-Quit orphan без ROT книги дополнительно проверен свой END с Close/Quit PASS.
Предложенный эксплуатационный порядок — такой же ownership/ROT review; при любых
сомнениях STOP и ручное решение. Автоматического убийства Excel в product нет.

## 8. RPC_E_DISCONNECTED HANDLING

Настоящий собственный Excel завершён непосредственно перед вызовом связанного
Save. Реальный HRESULT сохранён без переименования: текущая среда может вернуть
0x800706BA/0x800706BE либо 0x80010108. STOP безопасен, исходный XLSM/backup целы,
COMMITTED отсутствует, recovery = ABORTED_UNCHANGED. Отдельно точный 0x80010108
проверен через явно обозначенную Save fault injection на реальном собственном Excel.
Cleanup пробует все этапы и сохраняет исходную COM cause. Доказательства:
targeted_04/cases/real_COM_server_death_Save.json и RPC_E_DISCONNECTED_Save_injection.json.
Эти два вида проверки не выдаются друг за друга.
Причина первоначального спонтанного диагностического события из stage8_20261010_03
по-прежнему UNKNOWN; устранение первопричины не заявляется.

## 9. CRASH RECOVERY

Реальные worker os._exit(73): до Save, после native Save, перед publication,
посередине publication и после publication. Отдельно worker умер внутри настоящего
Save через Excel BeforeSave callback; test-only EnableEvents включён исключительно
на изолированной копии, callback marker доказывает активный вызов Save. Это fault
injection, не аппаратное испытание и не воспроизведение исходного спонтанного сбоя.
До записи: проверенный backup и неизменённый target, ABORTED без replay.
Полная запись + APPLYING: recovery только COMMITTED, registry_writes=0.
Частичная запись: STOP, никаких CLOSED/повторного списания/автоматического rollback.
Повреждённый backup запрещает публикацию; torn staged XLSM запрещает публикацию.

## 10. XLSM PRESERVATION

Native Excel normal reopen (CorruptLoad=0), формат 52, HasVBProject, disabled macros,
11 VBA streams побайтно идентичны. Формулы, стили, листы, veryHidden и package
structure сохранены; разрешены только C24/C25 и прежний узкий список save metadata
и formula caches. Native SUM после закрытия =23, =2+3 остаётся 5.
Копируются байты уже проверенного native-saved staged XLSM, package не пересобирается
Python. File identity и ACL исходного target сохраняются при обычном закрытии.
Evidence: native_normal_reopen_XLSM_preservation.json и transaction preservation.

## 11. TARGETED TEST RESULTS

Каноническая окончательная матрица: **38 PASS / 0 FAIL / 38 проверок**:
35 успешных сценариев targeted_03 + два COM-сценария targeted_04, все на одной
битово идентичной версии runtime. Дополнительно targeted_05 проверил native close,
backup/preservation и normal reopen после исправления инструкции; полный delivery SHA
совпал до/после этого завершающего теста. Повторный полный прогон 35 PASS не требовался,
код не менялся; доработаны только ошибочное ожидание HRESULT и отдельный
fault-injection тест точного HRESULT. CANONICAL_TARGETED_RESULTS.json явно
ссылается на заменённую failed assertion; исторический FAIL сохранён.
Все A–L covered; четырём прежним FAILED соответствует исправленная concurrency
группа (включая точный valid native 12/4/9) и cleanup_CutCopyMode.
Ограниченная регрессия касается close/state/backup/recover/native/office; Stage 1–7
не запускались. Полная таблица: CANONICAL_TARGETED_RESULTS.json; каждый протокол содержит
стенд, реальные process stdout/stderr, состояния журнала и diagnostics.

| Case | Result |
|---|---|
{test_table}

targeted_01 сохранён: 29 PASS / 5 FAIL; targeted_02: 27 PASS / 9 FAIL;
targeted_03: 35 PASS / 1 FAIL; targeted_04: 2 PASS / 0 FAIL; targeted_05: 1 PASS / 0 FAIL.
Единственное последующее изменение product — необходимая правка текста STAGE8_DEV.md;
после неё выполнен targeted_05 с новым fingerprint всей поставки.
Первый прогон выявил retained proxies и lazy Save name lookup masking HRESULT;
release/GC, освобождение локального proxy и раннее связывание Save исправлены.
Во втором прогоне overly-short 4-second exit timeout не учитывал подтверждённый
задержанный выход Excel (raw 2.41 sec; session 14.44 sec); ожидание увеличено до 25 sec.
Crash cleanup ожидал один START, хотя после staging и inspect могли жить два своих
post-Quit server. Теперь обработан каждый подтверждённый PID. HRESULT тест учитывает
фактически возвращённый RPC code; точный RPC_E_DISCONNECTED проверен отдельной
инъекцией. Последний historical FAIL targeted_03 — слишком узкое ожидание HRESULT,
а не незамеченный успешный debit: операция уже вернула STOP. Product между
targeted_02, targeted_03 и targeted_04 не менялся. Промежуточные FAIL не удалены.
Первичные smoke остановки на недоступных win32con constants устранены использованием
реальных win32file constants/документированного FILE_READ_ATTRIBUTES.

## 12. SOURCE INTEGRITY

Checkpoint до product edits: PRODUCT_CHECKPOINT.json + before_product/.
Состояние семи файлов при продолжении после UI interruption: RESUME_7_SHA.json.
Все **{integrity['protected_paths']}** protected paths совпали BEFORE/AFTER:
Stage 1–7, рабочие/DEV реестры, ВОМ, каталоги, templates, паспорта, прежние receipts
и validation outputs. Final evidence: final_integrity_02/SOURCE_INTEGRITY.json и
PROTECTED_BEFORE / final_integrity_02/PROTECTED_AFTER. Ошибок чтения/изменений 0.
Прежние canonical и independent product/evidence fingerprints сверены; Work08
manifest bindings совпали с checkpoint. Отдельного полного аудита источников нет.
Граница доказательства — перечисленные protected paths; неизвестные внешние
архивы не аттестованы.

Для новой тестируемой runtime версии:
**targeted_03/PRODUCT_BEFORE = PRODUCT_AFTER = PRODUCT_CURRENT**.
Те же SHA совпали с targeted_04/PRODUCT_BEFORE и PRODUCT_AFTER.
Те же runtime SHA совпали с targeted_05/PRODUCT_BEFORE и PRODUCT_AFTER.
Равенство проверялось после каждого сценария и при final seal. Старый product SHA
не используется как подтверждение новой версии. Новый manifest пишется последним.
FULL_PRODUCT_BINDING.json также связывает все {len(delivery_current)} файлов поставки:
**targeted_05/DELIVERY_BEFORE = DELIVERY_AFTER = CURRENT**, включая новую инструкцию.
Итоговый fingerprint: `{S.digest(delivery_current)}`.

## 13. CHANGED FILES / SHA

| Modified product | BEFORE SHA-256 | AFTER SHA-256 |
|---|---|---|
{sha_table}

| New targeted test file | SHA-256 |
|---|---|
{new_table}

SHA всех входных/выходных TEST registries, inputs, backups и journal entries:
TEST_INPUT_OUTPUT_JOURNAL_SHA.json. Полный пакет включает before reproduction,
промежуточный и окончательный прогоны; REVIEW_MANIFEST.json связывает каждый файл.

## 14. REMAINING LIMITATIONS

**Публикация через удерживаемый handle не является атомарной заменой всего файла.**
Обычный reader может увидеть partial bytes во время copy; смерть процесса посередине
может оставить непригодный XLSM. Такой файл не принимается за CLOSED: SHA не совпадает
ни с before, ни с after; остаётся APPLYING/STOP и проверенный backup. Автоматического
restore/replay нет. Это явная цена защиты namespace без release/reacquire race.
Полностью атомарная rename публикация при сохранении запрета DELETE не заявляется.
Backup/recovery proof проверен для process crashes, не физической потери питания.

os._exit/аппаратное завершение не гарантируют Excel cleanup. Контролируемая обработка
своих orphan является TEST procedure/предложением; product сам чужие PID не завершает.
Старые незавершённые операции без identity требуют review, не автоматического commit.
Локальный hash chain не защищает от администратора/пересчёта всех журналов.
Официальная выдача/реальные списания/production, ручной пользовательский E2E,
аппаратная потеря питания, иные файловые системы, Stage 1–7 full regression,
Stage 9, EXE, PDF/deployment — **NOT_RUN**. Причина исходного COM события UNKNOWN.

## 15. READY FOR TARGETED WORK 08 REVIEW: YES

Только новый bounded DEV/TEST fix и evidence, с явным review ограничения in-place
publication. Это готовность к независимой targeted проверке, не к официальной выдаче.
Реальных списаний 0. После отчёта STOP. Stage 9 не начат.
'''
(RUN/'FINAL_REPORT.md').write_text(report,encoding='utf8')
assert core.code_bindings()==before
manifest={'schema':'STAGE8_WORK08_FIX_MANIFEST_V1','created_utc':S.now(),
          'ready_for_targeted_work08_review':'YES','canonical_tests':str(RUN/'CANONICAL_TARGETED_RESULTS.json'),
          'product':delivery_current,'runtime':core.code_bindings(),
          'product_fingerprint_sha256':S.digest(delivery_current),'changed_product':changes,'new_tests':newtests,
          'evidence':{str(p):S.sha(p) for p in sorted(RUN.rglob('*')) if p.is_file() and p.name!='REVIEW_MANIFEST.json'}}
S.write(RUN/'REVIEW_MANIFEST.json',manifest)
# Readback only after the manifest-last publication.
for p,h in manifest['evidence'].items():assert S.sha(p)==h,p
assert core.code_bindings()==manifest['runtime']
assert {p:S.sha(p) for p in checkpoint}==manifest['product']
print('FINAL_SEAL_PASS',len(manifest['evidence']),'files',matrix['PASS'],'PASS',flush=True)
print('MANIFEST_SHA256',S.sha(RUN/'REVIEW_MANIFEST.json'),flush=True)
