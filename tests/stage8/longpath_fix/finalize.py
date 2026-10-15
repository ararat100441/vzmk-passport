"""Seal current source, fresh TEST data and canonical results; manifest last."""
import concurrent.futures, difflib, gc, json, sys
from pathlib import Path
from stage8_closure import native, office, storage as S
import win32api, win32con, win32process

RUN=S.ROOT/'validation/stage8_longpath_fix_20261011_001'
HERE=Path(__file__).parent
def write(p,q):S.write(p,q)
def alive(q):
    try:
        h=win32api.OpenProcess(win32con.PROCESS_QUERY_INFORMATION,False,q['pid'])
        try:return win32process.GetExitCodeProcess(h)==259 and win32process.GetProcessTimes(h)['CreationTime'].isoformat()==q['creation_time']
        finally:h.Close()
    except Exception:return False

if __name__=='__main__':
    current={str(p):S.sha(p) for p in [*sorted((S.ROOT/'stage8_closure').glob('*.*')),S.ROOT/'stage8.py',S.ROOT/'STAGE8_DEV.md',S.ROOT/'Запустить Этап 8 DEV.cmd'] if p.is_file()}
    assert current==S.read(RUN/'FINAL_TEST_PRODUCT_BEFORE.json')==S.read(RUN/'TEST_PRODUCT.json')
    roots=S.read(RUN/'TEST_ROOTS.json');long=Path(roots['long']);short=Path(roots['short'])
    cases=[]
    for batch in ['long_02','short_03','targeted_03']:
        q=S.read(RUN/batch/'RESULTS.json');rows=q['results'] if isinstance(q,dict) else q
        assert len(rows)==(25 if batch.startswith('targeted') else 27),(batch,len(rows))
        assert all(v['status']=='PASS' for v in rows),batch
        for v in rows:cases.append(v|{'batch':batch,'evidence':str(RUN/batch/'cases'/(v['name']+'.json'))})
    assert len(cases)==79
    with office.diagnostic_scope(RUN/'final_native_office'):
        before=S.read(RUN/'BEFORE_REPRODUCTION.json')
        source=Path(before['long']);data=S.sha(source)
        result=native.inspect(source,{'УЗК!C24':12,'УЗК!C25':10})
        assert data==before['sha256']==result['sha256']==S.sha(source)
        assert result['normal_open'] and result['read_only'] and result['macros_disabled']
        assert len(result['vba_streams'])==11 and result['formulas']['S8_FORMULAS!$A$1']['value']==30
        assert result['verified_copy']['source_and_copy_rechecked']
        assert not Path(result['verified_copy']['verification_path']).parent.exists()
        write(RUN/'FINAL_NATIVE_REOPEN.json',{'status':'PASS','source_length':len(str(source)),'native':result})
    gc.collect()
    cases.append({'name':'final_native_260_byte_exact_reopen','status':'PASS','batch':'final_native','evidence':str(RUN/'FINAL_NATIVE_REOPEN.json')})
    previous=S.read(S.ROOT/'validation/stage8_recovery_independent_20261011_01/RESULTS.json')
    failed=[v['name'] for v in previous['results'] if v['status']=='FAIL']
    canonical={v['name']:v for v in cases if v['batch']=='long_02'}
    assert len(failed)==15 and all(canonical[n]['status']=='PASS' for n in failed)
    write(RUN/'PREVIOUS_15_FAIL_RETEST.json',{'previous_PASS':12,'previous_FAIL':15,'independent_review_complete':False,'cases':[{'name':n,'before':'FAIL','after':'PASS','evidence':canonical[n]['evidence']} for n in failed]})
    write(RUN/'CANONICAL_RESULTS.json',{'PASS':80,'FAIL':0,'cases':cases,'visual_GUI_E2E':'NOT_RUN: native computer-control API unavailable; real Tk widget/handler integration PASS','product_fingerprint':S.digest(current)})
    old=S.read(RUN/'CHECKPOINT_PRODUCT.json');changes=[];diff=[]
    (RUN/'after_product').mkdir()
    for p,h in current.items():
        q=Path(p);(RUN/'after_product'/q.name).write_bytes(q.read_bytes())
        if old.get(p)==h:continue
        assert q.name in ('shortcopy.py','core.py','native.py','recovery.py'),p
        a=(RUN/'before_product'/q.name).read_text(encoding='utf8').splitlines(True) if p in old else []
        diff.extend(difflib.unified_diff(a,q.read_text(encoding='utf8').splitlines(True),fromfile='before/'+q.name,tofile='after/'+q.name))
        changes.append({'file':p,'before_sha256':old.get(p),'after_sha256':h})
    assert len(changes)==4
    (RUN/'EXACT_PRODUCT.diff').write_text(''.join(diff),encoding='utf8')
    write(RUN/'CHANGED_FILES_SHA.json',changes)
    write(RUN/'FINAL_TEST_PRODUCT_AFTER.json',current)
    write(RUN/'PRODUCT_CURRENT.json',current)
    write(RUN/'FULL_PRODUCT_BINDING.json',{'before_tests_equals_after_tests_equals_current':True,'fingerprint':S.digest(current),'delivery_files':len(current),'product':current})
    protected=S.read(RUN/'PROTECTED_BEFORE.json')
    def one(p):return p,S.sha(p)
    with concurrent.futures.ThreadPoolExecutor(max_workers=8) as ex:after=dict(ex.map(one,protected))
    assert after==protected,'PROTECTED_FILE_CHANGED'
    # The earlier read-only checkpoint is retained; this is the final after-tests gate.
    write(RUN/'PROTECTED_FINAL.json',after)
    write(RUN/'SOURCE_INTEGRITY_FINAL.json',{'status':'PASS','protected_files':len(after),'changes':[],'prior_review_not_modified':True})
    inventory=[]
    for base in [long,short]:
        for r in sorted((base/'sandboxes').iterdir()):
            if not r.is_dir():continue
            reg=r/'registry/УЗК_TEST.xlsm'
            item={'sandbox':str(r),'registry_sha256':S.sha(reg) if reg.exists() else None,'registry_identity':S.file_identity(reg) if reg.exists() else None,
                  'inputs_backups_journal_sha256':{str(p):S.sha(p) for sub in ['inputs','backups','journal'] for p in (r/sub).rglob('*') if p.is_file()}}
            try:
                h=S.history(r);item['journal']={'states':[v['state'] for v in h['events']],'closed_keys':list(h['closed']),'active':h['active']['state'] if h['active'] else None,'expected_identity':h['expected_identity'],'expected_sha256':h['expected_sha256']}
            except Exception as e:item['intentionally_faulted_journal_STOP']=str(e)
            inventory.append(item)
    write(RUN/'TEST_INPUT_OUTPUT_JOURNAL_SHA.json',inventory)
    records=[]
    for base in [RUN,long,short]:
        for p in base.rglob('START.json'):
            q=S.read(p)
            if q.get('schema')=='STAGE8_OWNED_EXCEL_V1':records.append((p,q))
    live=[{'ownership_evidence':str(p),**q} for p,q in records if q.get('owned') and alive(q)]
    write(RUN/'FINAL_EXCEL_OWNERSHIP.json',{'owned_records':sum(bool(q.get('owned')) for _,q in records),'own_live':live,'own_live_count':len(live),'finalizer_terminated_processes':False})
    assert not live,'OWN_EXCEL_LIVE'
    # Keep only workspace paths actually identified in this task's ownership
    # evidence. Never scan or purge the user's generic temporary directories.
    workspaces={Path(n).parent for _,q in records for n in q.get('expected_workbooks',[]) if Path(n).parent.name.startswith('s8n_')}
    leftovers=[]
    for w in sorted(workspaces):
        if w.exists():leftovers.append({'path':str(w),'files':{str(p):S.sha(p) for p in w.iterdir() if p.is_file()},'automatic_purge':False})
    write(RUN/'CONTROLLED_CRASH_WORKSPACE_LEFTOVERS.json',leftovers)
    wb=[(v['batch'],v['name'],v['evidence']) for v in cases]
    matrix='\n'.join('| '+b+' | '+n+' | PASS | '+str(Path(e).relative_to(RUN))+' |' for b,n,e in wb)
    changed='\n'.join('| '+Path(v['file']).name+' | `'+(v['before_sha256'] or 'NEW')+'` | `'+v['after_sha256']+'` |' for v in changes)
    fifteen='\n'.join('| '+n+' | FAIL → PASS |' for n in failed)
    source=S.ROOT/'validation/stage8_recovery_independent_20261011_01'
    windows=sys.getwindowsversion()
    write(RUN/'ENVIRONMENT.json',{'python':sys.version,'windows_build':windows.build,'storage':'fixed local NTFS','current_user_LOCALAPPDATA':str(__import__('stage8_closure.shortcopy',fromlist=['BASE']).BASE),'packaged_child_LOCALAPPDATA':'Windows can redirect to Packages/OpenAI.Codex.../LocalCache/Local; actual copy paths are in each native proof','native_GUI_control':'unavailable; no workaround used','WMI_used':False})
    report=f'''# ВЗМК — Этап 8 — TARGETED FIX — LONG PATH NATIVE EXCEL

11.10.2026, Europe/Samara. Только новые искусственные TEST XLSM.
**80 PASS / 0 FAIL в окончательной матрице. READY FOR TARGETED WORK 08 REVIEW: YES.**
Это авторская targeted validation. Прерванный Escape независимый Work 08 не завершён и не объявляется PASS.

## 1. ROOT CAUSE
В recovery.py initial-stage native.inspect получал исходный restore_<id>/verified.xlsm длиной 260 символов.
Python/NTFS читали книгу, но установленный Excel 16 Workbooks.Open возвращал 0x80020009 / -2146827284.
В этой сессии до исправления новые копии одного SHA `{before['sha256']}` дали длинный STOP / короткий PASS.
Длина 260 — воспроизведённый путь конкретного Excel, не универсальный лимит Windows.
Исходные PATH_LENGTH_PROOF.json, GUI_VALIDATION.json, RESULTS.json и предыдущий FINAL_REPORT.md прочитаны; их SHA связаны protected checkpoint.

## 2. EXACT FIX / DIFF
EXACT_PRODUCT.diff содержит полный diff; before_product / after_product — точные снимки.
shortcopy.py добавляет безопасный единый маршрут native-проверки. native.py направляет read-only inspect и частное native Save через короткую папку.
core.py выполняет настоящий native preflight до INTENT/backup/изменения реестра и сохраняет ту же папку до конца операции.
recovery.py проверяет backup тем же маршрутом до любых rename, включая resume; конечные байты читает из уже удерживаемого RegistryGuard.
Writer stage также перенесён в эту папку: проверка одного restore-пути оставила бы отдельный длинный Workbooks.Open в close.
storage.py, office.py, GUI и CLI не изменялись. Протокол восстановления и его состояния сохранены.

## 3. SHORT VERIFIED COPY SAFETY
Каждая операция создаёт уникальный LOCALAPPDATA/s8n_<32 hex> через Win32 CreateDirectory с protected DACL: только текущий пользователь и SYSTEM, наследование на файлы.
Namespace handles удерживают локальные NTFS-родители без DELETE sharing. Создание копии — exclusive, flush/fsync; общий временный файл не используется.
Источник удерживается frozen+namespace guard либо читается через уже существующий write/delete RegistryGuard без ослабления sharing.
Копия открыта GENERIC_READ / FILE_SHARE_READ: внешние WRITE/DELETE/replace запрещены до последней SHA/identity проверки после Excel.
Evidence содержит исходный путь, identity, SHA, уникальный короткий путь, его identity и SHA, exact_bytes и source_and_copy_rechecked.
Очистка использует DELETE handle, сравнивает identity и удаляет именно этот объект через FileDispositionInfo; чужая подмена сохраняется и вызывает STOP.
Обычное исключение очищает файл и каталог, исходное исключение сохраняется; ошибки очистки добавляются в notes.
Политика приложения ограничивает native-путь 180 символами; фактический Open проверяется, а не выводится из длины.

Win32 sharing: [CreateFileW](https://learn.microsoft.com/en-us/windows/win32/api/fileapi/nf-fileapi-createfilew).
Удаление через handle: [SetFileInformationByHandle](https://learn.microsoft.com/en-us/windows/win32/api/fileapi/nf-fileapi-setfileinformationbyhandle).
Normal load без recovery через object model: [Workbooks.Open](https://learn.microsoft.com/en-us/office/vba/api/excel.workbooks.open).

## 4. LONG PATH BEFORE/AFTER
BEFORE_REPRODUCTION.json: исходный путь 260, STOP; те же байты по короткому пути PASS.
FINAL_NATIVE_REOPEN.json: тот же исходный путь и SHA — PASS через защищённую копию, файл не менялся.
long_02 воспроизводит длину исходного стенда: restore verified.xlsm снова 260 символов, все 27 PASS.
Дополнительный native source-path 370 символов также PASS (это проверка установленной среды, не обещание любой длины).

## 5. PREVIOUS 15 FAIL RESULTS
PREVIOUS_15_FAIL_RETEST.json связывает все имена исходных FAIL с fresh long_02 evidence.

| Сценарий | Результат |
|---|---|
{fifteen}

## 6. SHORT PATH REGRESSION
short_03: все 27 исходных сценариев PASS. Функции оригинального review сохранены;
единственная адаптация его копии — NATIVE_EXACT/POSITIVE пишутся в отдельный batch, чтобы прогоны не перезаписывали evidence друг друга.
baseline_helpers.py скопирован побайтно; HARNESS_BINDINGS.json фиксирует SHA источника и копий. Старые outputs не записывались.

## 7. NATIVE EXCEL / VBA CHECK
Настоящий Microsoft Excel 16, FileFormat=52, ReadOnly=True, AutomationSecurity=3, UpdateLinks=0, CorruptLoad=0, без Save источника.
Проверены все 11 VBA streams, все ZIP parts восстановления побайтно, формулы и результаты 30/5, стили, sheets/visibility, definedNames и relationships.
После нового подтверждённого close меняются только разрешённые C24/C25 и ранее разрешённые save metadata/caches; полное сравнение package сохранено.
Финальный native reopen отдельно PASS; копия и её каталог удалены.

## 8. BACKUP / JOURNAL SAFETY
Частичная XLSM после реального os._exit в during_publish обнаруживается; обычный recover/close STOP без replay.
Restore-preview проверяет exact backup BEFORE/SHA/журнал; подтверждённый restore заканчивается RECOVERED_ABORTED, decrements=0, closed=False, без COMMITTED.
Старая и новая identity связаны recovery operation; неизвестная внешняя identity по-прежнему STOP.
Новый процесс и новый подтверждённый preview/close дают ровно одно списание 12→9 и 10→6, один COMMITTED; повтор ALREADY_CLOSED_TEST ничего не пишет.
Проверены сохранение более поздних COMMITTED, backup damage/substitution, journal damage, namespace-gap creator, native replace до/после публикации, parent rename.
Недоступная/неверная/слишком длинная/запрещённая ACL short-base даёт STOP до INTENT или target rename; source/backup/journal не меняются.
Crash restore проверен в шести точках с новым процессом, fresh preview и explicit resume без автоматического rollback/debit.

## 9. GUI / CLI
CLI: реальные отдельные процессы restore-preview, отмена, подтверждённый restore на длинном пути PASS.
GUI: реальные Tk widgets и handlers, default NO, cancel unchanged, confirmed restore, RECOVERED_ABORTED/0 списаний, новый App/preview, short-path failure STOP — PASS.
В GUI integration тестах messagebox ответы заданы тестом; ручной/видимый Computer Use E2E — NOT_RUN: native control API недоступен в текущей сессии. Ограничение не обходилось.
Предыдущая видимая GUI_VALIDATION — историческое evidence, не новый GUI PASS и не завершение независимого review.

## 10. TARGETED TESTS
Окончательные прогоны: long_02 27 + short_03 27 + targeted_03 25 + финальный native reopen 1 = 80 PASS.
Дополнительно проверены отдельные внешние процессы против source/copy, SHA/identity подмена до Excel, normal exception cleanup, два одновременных close,
существующий writer/writable mapping, torn native-save stage, реальная смерть TEST-процесса с открытой Excel-книгой, backup и safe abort без COMMITTED.
MAPPING_LOCK_PROBE.json дополнительно подтверждает: live writable mapping блокирует frozen acquisition даже после закрытия исходного writer handle.
RPC_E_DISCONNECTED 0x80010108 injection и настоящий COM disconnect при Save сохраняют safe STOP и диагностику. Прежняя спонтанная первопричина остаётся UNKNOWN.
DEV 9/2026 остаётся BLOCKED; официального закрытия/реального списания нет.

Предварительные прогоны сохранены отдельно и не используются для PASS окончательного продукта:
long_01/short_01: 27/27 на промежуточном SHA до исправления очистки пустого каталога;
targeted_01: 21 PASS/4 FAIL (один дефект очистки; три ошибки тестового стенда: отсутствующая константа, проверка длины, redirected LOCALAPPDATA);
targeted_02: 24 PASS/1 FAIL (тест ожидал другой текст безопасного ACL STOP).
short_02: 26 PASS/1 FAIL — при очистке собственного тестового orphan ошибочно завершён собственный Excel активного TEST-прогона.
Его 0x800706be STOP и проверенный backup сохранены; прогон целиком исключён и полностью повторён как short_03.
PRELIMINARY_OWN_ORPHAN_CLEANUP.json сохраняет фактические PID/creation/evidence; его первоначальное пояснение аварийности для PID 70112 было ошибочным.
После этого никаких очисток процессов во время тестов не выполнялось. Чужой/historical Excel не завершался.

| Batch | Сценарий | Результат | Evidence |
|---|---|---|---|
{matrix}

## 11. SOURCE INTEGRITY
Checkpoint сделан до правок. `{len(after)}` защищённых файлов: BEFORE = FINAL, изменений 0.
Защищены Stage 1–7, реальные/DEV реестры, ВОМ, справочники, шаблоны, документы, прежние receipts и все прежние validation/test outputs.
PRODUCT BEFORE TESTS = AFTER TESTS = CURRENT для 14 файлов; fingerprint `{S.digest(current)}`.
SOURCE_INTEGRITY_FINAL.json, FULL_PRODUCT_BINDING.json и TEST_INPUT_OUTPUT_JOURNAL_SHA.json содержат SHA, identity и состояния журналов новых TEST-копий.
Полная регрессия Stage 1–7 не запускалась. EXE/PDF/deployment не выполнялись. Stage 9 не начат.

## 12. CHANGED FILES / SHA

| Файл Stage 8 | До | После |
|---|---|---|
{changed}

Новые targeted tests: tests/stage8/longpath_fix/targeted.py и finalize.py; SHA перечислены в REVIEW_MANIFEST.json.
Точные before/after snapshots и diff включены в новый manifest; старые manifests не переписывались.

## 13. REMAINING LIMITATIONS
Только fixed local NTFS и доступный приватный LOCALAPPDATA с проходящим реальным Excel preflight. При иных условиях STOP без небезопасного fallback.
Политика 180 символов консервативна; native Excel разных версий может иметь другие ограничения.
При os._exit/аппаратной смерти finally не гарантируется: собственный Excel или private workspace могут остаться.
Тестовая очистка использовала только собственные PID+creation time+ownership evidence; продукт не содержит global taskkill или автоматического завершения чужих процессов.
Идентифицированные crash leftovers сохранены в CONTROLLED_CRASH_WORKSPACE_LEFTOVERS.json, не переиспользуются и автоматически не удаляются.
Неизвестный дополнительный файл/изменённая identity в workspace приводит к STOP с notes и сохранению данных, а не рекурсивному удалению.
COM/диск могут отказать после preflight: остаётся прежний безопасный pending/recovery протокол, без ложного CLOSED и повторного списания.
Видимый GUI E2E NOT_RUN; functional Tk integration и native Excel COM выполнены. Спонтанный исторический 0x80010108 UNKNOWN.

## 14. READY FOR TARGETED WORK 08 REVIEW
**YES.** Исправление и окончательные targeted checks готовы к новой независимой проверке.
Это не утверждение, что прерванный независимый review завершился PASS.

STOP. Этап 9 не начат.
'''
    (RUN/'FINAL_REPORT.md').write_text(report,encoding='utf8')
    assert {p:S.sha(p) for p in current}==current
    tests={str(p):S.sha(p) for p in HERE.glob('*.py')}
    evidence={str(p):S.sha(p) for base in [RUN,long,short] for p in base.rglob('*') if p.is_file() and '__pycache__' not in p.parts and p.name!='REVIEW_MANIFEST.json'}
    manifest={'schema':'STAGE8_LONGPATH_FIX_REVIEW_V1','product':current,'product_fingerprint':S.digest(current),'tests':tests,'evidence':evidence,'canonical_PASS':80,'canonical_FAIL':0,'protected_files_unchanged':len(after),'independent_review_complete':False,'ready_for_targeted_review':True}
    write(RUN/'REVIEW_MANIFEST.json',manifest)
    # Readback also checks that the sealed package is not missing a referenced file.
    for section in ['product','tests','evidence']:
        assert all(S.sha(p)==h for p,h in manifest[section].items()),section
    print('SEALED',len(evidence),'80 PASS','manifest_sha256',S.sha(RUN/'REVIEW_MANIFEST.json'),'product',S.digest(current),flush=True)
