"""Read-only final verification, exact diff, full evidence binding; manifest last."""
import difflib,gc,json,platform,time,threading
from pathlib import Path
from stage8_closure import storage as S
import win32api,win32con,win32process
RUN=S.ROOT/'validation/stage8_recovery_fix_20261011_01'
HERE=Path(__file__).parent
def write(path,value):
    if Path(path).exists():
        assert S.read(path)==value,('EXISTING_SEAL_OUTPUT_DIFFERS',path)
    else:S.write(path,value)
def live(record):
    try:
        h=win32api.OpenProcess(win32con.PROCESS_QUERY_INFORMATION,False,record['pid'])
        try:return win32process.GetExitCodeProcess(h)==259 and win32process.GetProcessTimes(h)['CreationTime'].isoformat()==record['creation_time']
        finally:h.Close()
    except Exception:return False

if __name__=='__main__':
    before=S.read(RUN/'PRODUCT_BEFORE.json')
    paths=[*sorted((S.ROOT/'stage8_closure').glob('*.*')),S.ROOT/'stage8.py',S.ROOT/'STAGE8_DEV.md',S.ROOT/'Запустить Этап 8 DEV.cmd']
    product={str(p):S.sha(p) for p in paths if p.is_file()}
    batches=['targeted_04','boundaries_02','regression_01'];results=[]
    for batch in batches:
        assert S.read(RUN/batch/'PRODUCT_BEFORE.json')==S.read(RUN/batch/'PRODUCT_AFTER.json')==product,batch
        for r in S.read(RUN/batch/'RESULTS.json'):
            assert r['status']=='PASS',r
            results.append(r|{'source':str(RUN/batch/'cases'/(r['case']+'.json'))})
    assert len(results)==50,len(results)
    write(RUN/'CANONICAL_RESULTS.json',{'PASS':50,'FAIL':0,'product_fingerprint':S.digest(product),'cases':results})
    write(RUN/'PRODUCT_CURRENT.json',product)
    write(RUN/'FULL_PRODUCT_BINDING.json',{'status':'PASS','before_tests_equals_after_tests_equals_current':True,
            'batches':batches,'product':product,'fingerprint':S.digest(product),'delivery_files':len(product)})
    (RUN/'after_product').mkdir(exist_ok=True)
    changes=[];diff=[]
    for p,h in product.items():
        old=before.get(p);q=Path(p);(RUN/'after_product'/q.name).write_bytes(q.read_bytes())
        if old==h:continue
        assert q.name in ('core.py','storage.py','gui.py','stage8.py','STAGE8_DEV.md','recovery.py'),p
        a=(RUN/'before_product'/q.name).read_text(encoding='utf8').splitlines(True) if old else []
        b=q.read_text(encoding='utf8').splitlines(True)
        diff.extend(difflib.unified_diff(a,b,fromfile='before/'+q.name,tofile='after/'+q.name))
        changes.append({'path':p,'before_sha256':old,'after_sha256':h})
    (RUN/'EXACT_PRODUCT.diff').write_text(''.join(diff),encoding='utf8')
    write(RUN/'CHANGED_FILES_SHA.json',changes)
    protected=S.read(RUN/'PROTECTED_BEFORE.json');after={};errors={}
    monitor_stop=threading.Event();progress={'index':0,'path':None}
    def monitor():
        while not monitor_stop.wait(10):print('HASH_CURRENT',progress,flush=True)
    threading.Thread(target=monitor,daemon=True).start()
    for i,(p,h) in enumerate(protected.items()):
        progress.update(index=i,path=p)
        try:
            after[p]=S.sha(p)
            if after[p]!=h:errors[p]='SHA_CHANGED'
        except Exception as e:errors[p]=str(e)
        if i%5000==0:print('PROTECTED',i,flush=True)
    monitor_stop.set()
    S.write(RUN/'PROTECTED_AFTER.json',after)
    S.write(RUN/'SOURCE_INTEGRITY.json',{'status':'PASS' if not errors else 'FAIL','protected_files':len(protected),
            'errors':errors,'modified_existing_product':sum(c['before_sha256'] is not None for c in changes),
            'new_product':sum(c['before_sha256'] is None for c in changes),'office_native_unchanged':all(product[p]==h for p,h in before.items() if Path(p).name in ('office.py','native.py'))})
    assert not errors,errors
    records=[(p,S.read(p)) for p in RUN.rglob('START.json')]
    own=[{'start':str(p),**r} for p,r in records if r.get('owned') and live(r)]
    S.write(RUN/'FINAL_EXCEL_OWNERSHIP.json',{'owned_records':sum(r.get('owned',False) for _,r in records),
            'own_live':own,'own_live_count':len(own),'process_termination_by_finalizer':False})
    assert not own,'OWNED_EXCEL_STILL_LIVE'
    datasets=[]
    for r in sorted((RUN/'sandboxes').iterdir()):
        target=r/'registry/УЗК_TEST.xlsm'
        entry={'sandbox':str(r),'registry_sha256':S.sha(target) if target.exists() else None,
               'registry_identity':S.file_identity(target) if target.exists() else None,
               'files':{str(p):S.sha(p) for sub in ('inputs','backups','journal') for p in (r/sub).rglob('*') if p.is_file()}}
        try:
            h=S.history(r);entry['journal']={'active':h['active']['state'] if h['active'] else None,
                'events':[e['state'] for e in h['events']],'closed':list(h['closed']),'expected_sha256':h['expected_sha256'],
                'expected_identity':h['expected_identity']}
        except Exception as e:entry['intentional_damaged_fixture_journal_stop']=str(e)
        datasets.append(entry)
    S.write(RUN/'TEST_INPUT_OUTPUT_JOURNAL_SHA.json',datasets)
    tests={str(p):S.sha(p) for p in sorted(HERE.glob('*.py'))}
    S.write(RUN/'NEW_TEST_FILES_SHA.json',tests)
    S.write(RUN/'ENVIRONMENT.json',{'python':platform.python_version(),'windows':platform.platform(),
            'storage_scope':'fixed local NTFS','date':'2026-10-11','timezone':'Europe/Samara'})
    table='\n'.join('| '+Path(c['path']).name+' | `'+(c['before_sha256'] or 'NEW')+'` | `'+c['after_sha256']+'` |' for c in changes)
    matrix='\n'.join('| '+r['case']+' | PASS | '+str(Path(r['source']).relative_to(RUN))+' |' for r in results)
    report=f'''# ВЗМК — Этап 8 — FINAL TARGETED FIX — XLSM RECOVERY

11.10.2026, Europe/Samara. Только новые искусственные TEST XLSM.
**50 PASS / 0 FAIL в окончательной матрице. READY FOR TARGETED WORK 08 REVIEW: YES.**
Это готовность к независимой targeted проверке, не её результат и не разрешение официальной выдачи.

## 1. ROOT CAUSE

Независимый Work 08: 34 PASS / 1 FAIL, MAJOR 1. Прежний recover умел лишь ABORTED
для BEFORE и COMMITTED без записи для полного AFTER. Для частичного XLSM он правильно
делал STOP, но поставляемой операторской команды возврата в работу не было. Ручная
замена проверенным backup меняла identity и оставалась запрещённой. До исправления это
повторено реальной смертью worker во время publication на новом стенде: BEFORE_REPRODUCTION.json.
Исходные reports/manifests и их evidence проверены до изменений: PRIOR_BINDINGS.json.

## 2. EXACT CODE CHANGES

Новый recovery.py: отдельные sealed preview, текст последствий, точная confirmation,
backup/native проверки, защищённые namespace transitions, resume и RECOVERED_ABORTED.
storage.py: дополняет append-only grammar и expected identity разрешённым restore;
добавляет non-replacing handle rename и directory namespace guard. core.py использует
разрешённую журналом identity и запрещает обычному recover автоматически продолжать restore.
gui.py добавляет состояние аварии и отдельную кнопку/подтверждение. stage8.py доставляет
CLI restore-preview/restore. STAGE8_DEV.md содержит полный операторский порядок.
office.py и native.py **побайтно неизменны**. Прежний target READ/WRITE handle без
WRITE/DELETE sharing и удержание до COMMITTED сохранены; их регрессия повторена.
Необходимость изменения directory helper доказана отдельными NTFS probes (раздел 7).
Полный exact diff: EXACT_PRODUCT.diff; before_product/ и after_product/ — exact snapshots.

## 3. OPERATOR RECOVERY FLOW

В том же TEST-стенде выбрать прежний inputs/CLOSE_REQUEST.json. UI показывает
«Обнаружена незавершённая операция», статус/причину, проверенный backup/SHA, реестр,
доступность и последствия. Закрытие заблокировано. Нажать «Восстановить из резервной копии»;
отдельное окно с default=Нет повторяет сведения. Отмена не меняет журнал/реестр.
После успешной native сверки: RECOVERED_ABORTED, паспорт не CLOSED, списаний 0.
Для закрытия заново выбрать запрос, получить новый preview, проверить входы и подтвердить
новую операцию. Старый close preview и старый restore preview не дают повторить действие.
CLI: restore-preview <sandbox> --save-preview <новое_имя.json>, затем restore <preview>
--confirm <точная строка>. Preview-файлы не перезаписываются; для повторного просмотра
используется новое имя. Обычный recover по-прежнему только завершает/отменяет журнал.

## 4. BACKUP VALIDATION

Полная цепь/sequence/seals, активная APPLYING операция, не COMMITTED, ключ не закрыт;
BEFORE равен ожидаемому журналом состоянию после прежних закрытий. Точные пути registry,
backups/<исходная операция>/УЗК_TEST.xlsm и SHA256.json, SHA обоих, verified=true,
принадлежность исходному plan и native STAGED/preservation проверяются под handles.
В исходном аварийном состоянии прежняя identity обязательна, SHA отличается от BEFORE
и AFTER, XLSM действительно повреждена/CRC failure. Целая неизвестная книга — STOP.
Восстановление идёт только из точного backup; Excel нормально открывает проверяемые
байты read-only без VBA/events/link updates. Тесты damaged/substituted backup, journal,
private stage и восстановленного native evidence дают STOP без изменения target.
SHA-chain защищает от повреждения/случайной подмены, не от администратора с пересчётом всей цепи.

## 5. FILE IDENTITY RECONCILIATION

До изменения namespace durable RESTORE_PREPARED фиксирует recovery id, исходную операцию,
backup, повреждённый SHA, old identity, new identity заранее созданной проверенной копии,
её путь и quarantine. Проверяются volume serial + file index. RECOVERED_ABORTED связывает
verified restored SHA/identity и native evidence; history принимает новую identity только
по этой законной цепи. Старые events/receipts не переписываются. Новая identity реально
отличалась в положительных тестах; fresh-process recover/preview приняли её. Неизвестная
замена после восстановления, даже byte-identical, снова STOP. Ручная замена backup до
официального restore также остаётся STOP — обход identity не добавлен.

## 6. JOURNAL STATE SAFETY

A: COMMITTED → restore forbidden, старый backup не отменяет списание.
B: полный AFTER при APPLYING → restore запрещён; прежний recover дописывает COMMITTED,
registry_writes=0. C: повреждённая частичная запись → явный restore:
APPLYING → RESTORE_PREPARED → RESTORE_APPLYING → RECOVERED_ABORTED. Ключ не добавляется
в closed, decrements=0. Новое закрытие — отдельный INTENT с новым id/preview/confirmation.
Обычный recover при незавершённом restore выдаёт STOP и требует операторское подтверждение.
Более поздние успешные закрытия защищены grammar/BEFORE и проверкой отсутствия active
операции; старый restore после двух новых COMMITTED отклонён. Восстановление следующей
частичной операции сохранило оба прежних закрытия и остаток 6, а не старые 12.

## 7. CRASH DURING RESTORE

Backup копируется/flush/fsync в приватный stage на том же томе и проходит native Excel
до target mutation. Source/target handles имеют WRITE/DELETE access для собственных
действий, но share только READ; внешние writer/delete/replace запрещены. Rename выполняется
через SetFileInformationByHandle(FileRenameInfo), ReplaceIfExists=FALSE, без fallback.
Первый rename сохраняет damaged.xlsm, второй публикует целый verified.xlsm. Между ними
target кратко отсутствует: это **не атомарная транзакция двух rename**. Внешний creator
может занять имя; тогда второй rename STOP, чужой файл сохранён, backup/stage/quarantine
сохранены, успех/COMMITTED отсутствуют. Такое состояние требует расследования, не auto restore.

Directory guard использует FILE_LIST_DIRECTORY|FILE_READ_ATTRIBUTES, READ|WRITE sharing,
без DELETE sharing. Attribute-only handle не блокировал rename каталога без target;
это независимо доказано DIRECTORY_SHARING_PROBE.json. LIST_DIRECTORY с одним READ
sharing блокировал и собственный child rename; DIRECTORY_RENAME_PROTOCOL_PROBE.json
подтвердил рабочее сочетание с WRITE sharing и сохранённым запретом directory rename.
Target handles по-прежнему не разрешают WRITE/DELETE sharing. В финальном тесте внешний
процесс не смог переименовать registry parent даже при отсутствующем target.

Реальные os._exit(73) проверены посреди private copy, после intent, перед quarantine,
после quarantine, после publication и перед terminal event. Новый процесс/preview
различает ровно TORN, PREPARED, VACANT и PUBLISHED по заранее записанным identities/SHA;
для продолжения нужно новое подтверждение. Никаких операций с неизвестными файлами.
Смерть во время private copy оставляет только непригодный приватный stage, не target.
Завершение требует повторной native сверки точных восстановленных байтов. Это proof
process-crash на local NTFS; физическая потеря питания не проверена.

Microsoft API: [CreateFile sharing](https://learn.microsoft.com/en-us/windows/win32/api/fileapi/nf-fileapi-createfilew),
[FILE_RENAME_INFO no-replace](https://learn.microsoft.com/en-us/windows/win32/api/winbase/ns-winbase-file_rename_info),
[SetFileInformationByHandle](https://learn.microsoft.com/en-us/windows/win32/api/fileapi/nf-fileapi-setfileinformationbyhandle).
Практические свойства данной Windows/NTFS подтверждены собственными probes и real-process tests.

## 8. NO DOUBLE DECREMENT

Restore никогда не вызывает native.stage и не применяет deductions. Восстановленные
байты строго BEFORE, после restore 12/10, журнал не CLOSED/COMMITTED. Повтор restore,
старый close preview, новый запуск и параллельные restore не списали ничего. Новый
явно подтверждённый close списал ровно один раз до 9/6; его повтор вернул ALREADY_CLOSED_TEST.
Два simultaneous close worker также сохранили ровно один COMMITTED.

## 9. XLSM PRESERVATION

Реестр после restore побайтно равен backup. Native normal open (CorruptLoad=0),
format52, read-only, HasVBProject подтверждены. Все 11 VBA streams идентичны; полный
нормализованный OOXML package, формулы, стили, листы/visibility совпали с backup.
Исходные 12/10, SUM=30, =2+3=5 восстановлены; новый close даёт 9/6 и SUM=23,
формулы сохраняются. Для финальной проверки при retained DELETE handle native Excel
читает приватную копию точных target bytes с тем же SHA; handles не отпускаются.
Дополнительно после выхода из guard сам target нормально открыт native Excel.

## 10. GUI VALIDATION

Настоящие Tk widgets в withdrawn test window: загрузка partial, текст всех обязательных
полей, close disabled, restore enabled, отдельное подтверждение/default=no, отмена без
write, действительное restore/native validation, RECOVERED_ABORTED/0 списаний, сброс
старых preview/confirmation. Новый App после restore получил новый close preview без
подтверждения. Повреждённый backup: точный BACKUP_DAMAGED, обе кнопки записи disabled.
CLI separate preview/confirm/restore выполнен реальными отдельными процессами.
Ручной визуальный пользовательский E2E — NOT_RUN; программная проверка не объявлена ручной.

## 11. TARGETED TEST RESULTS

**26 recovery acceptance + 6 boundary + 18 previous-fix regression = 50 PASS / 0 FAIL.**
Все финальные batches проверили полную поставку: BEFORE TESTS = AFTER TESTS = CURRENT,
включая doc и новый recovery.py. Fingerprint: `{S.digest(product)}`.

| Сценарий | Итог | Evidence |
|---|---|---|
{matrix}

История сохранена: targeted_01 25 PASS/1 FAIL (третий искусственный close запросил 4
S8-B при остатке2; исправлен только TEST-input); targeted_02 1 PASS. boundaries_01
5 PASS/1 FAIL выявил недостаточный directory attribute-only guard. targeted_03 с
LIST_DIRECTORY/READ sharing остановлен после 6 FAIL безопасного STOP: режим также
запрещал собственный child rename; остальные 20 сценариев этого прогона NOT_RUN.
Две промежуточные product snapshots и probes сохранены. Эти результаты не включены
в окончательные PASS. Финальные targeted_04, boundaries_02, regression_01 все связаны
с одним текущим fingerprint; product во время них не менялся. Первичный smoke поймал
длинный Excel path, исправление короткой verification copy затем проверено всеми batches.

## 12. PREVIOUS WORK 08 FAILURE

BEFORE_REPRODUCTION.json: реальный partial publication, свежий recover STOP; обычная
замена backup вернула12/10, но изменила identity и оставила pending. После fix поддерживаемый
операторский путь на свежем partial fixture завершил RECOVERED_ABORTED, новая identity
принята новым процессом, байты/backup сверены, новый preview/close списал один раз.
Финальный case partial_XLSM_supported_operator_recovery — PASS. Неизвестная ручная замена
по-прежнему запрещена; защищённые старые independent fixtures не исправлялись и не менялись.
Четыре прежних concurrent-write/cleanup FAILED сценария и связанные critical-phase,
параллельные close, mapping, первичное Save exception, чужой TEST Excel и HRESULT проверки
повторены на текущем коде и PASS. Exact0x80010108 — явная fault injection на настоящем Excel;
real Excel.Quit disconnect проверен отдельно. Первопричина прежнего спонтанного события UNKNOWN.

## 13. SOURCE INTEGRITY

Checkpoint выполнен до edits. Все **{len(protected)}** защищённых файлов неизменны,
ошибок чтения/изменений0: SOURCE_INTEGRITY.json и PROTECTED_BEFORE/AFTER.json.
Stage1–7, реальные/DEV реестры, документы/шаблоны/справочники, прежние receipts и все
предыдущие validation outputs не изменены. PRIOR_BINDINGS проверил прежний fix и
independent manifests/evidence до изменения product. Полный аудит Stage1–7 не повторялся.
Индекс вернул INTERNAL_ERROR; использованы точные исходники и bounded reads.
FINAL_EXCEL_OWNERSHIP.json: собственных живых Excel0; финализатор не завершал процессы.
TEST_INPUT_OUTPUT_JOURNAL_SHA.json связывает TEST inputs, final registry SHA/identity,
backup/metadata и journal events, включая сохранённые исторические failed fixtures.

## 14. CHANGED FILES / SHA

| Product | BEFORE SHA-256 | AFTER SHA-256 |
|---|---|---|
{table}

Новых test helpers: {len(tests)}; каждый SHA — NEW_TEST_FILES_SHA.json. Включены точные
source snapshots, diff, все входы/выходы/журналы и окончательные матрицы. REVIEW_MANIFEST.json
пишется последним и связывает каждый файл нового evidence-пакета, current product и tests.
Старый product SHA не используется как доказательство новой версии.

## 15. REMAINING LIMITATIONS

Исходное close publication остаётся in-place copy: смерть процесса может оставить
partial XLSM. Теперь есть явный проверенный операторский restore; автоматических rollback
или повторных списаний нет. Два restore rename не атомарны вместе: кратко отсутствующий
target и чужой creator дают предусмотренный STOP без потери чужих байтов. Unknown identity,
целая неизвестная книга, повреждённый journal/backup/evidence требуют расследования.
Локальная SHA-chain не является электронной подписью. Новый файл наследует ACL private
recovery directory; product остаётся только искусственным TEST, не production adapter.
Native Excel/path/COM failure может оставить restore pending; никаких SUCCESS до final
readback/native proof. os._exit не гарантирует finally/Excel cleanup; обработка own orphan
требует доказанного ownership. Чужие Excel и global taskkill не использовались.
Не проверены: аппаратное отключение питания, сеть/другие FS, администратор/злонамеренный
полный пересчёт evidence, ручной визуальный userE2E, official/production/real debit,
полная Stage1–7 regression, EXE/PDF/deployment — **NOT_RUN**. Root cause спонтанного COM
disconnect остаётся UNKNOWN. Stage9 не запускался.

## 16. READY FOR TARGETED WORK 08 REVIEW: YES

Только новый bounded DEV/TEST recovery fix с полной итоговой проверкой и явными
ограничениями. Независимая приёмка ещё не выполнена. Реальных списаний0. После отчёта STOP.
'''
    (RUN/'FINAL_REPORT.md').write_text(report,encoding='utf8')
    manifest={'schema':'STAGE8_RECOVERY_FIX_EVIDENCE_V1','time':S.now(),'product':product,'tests':tests,
              'product_fingerprint':S.digest(product),'PASS':50,'FAIL':0,
              'evidence':{str(p):S.sha(p) for p in sorted(RUN.rglob('*')) if p.is_file() and p.name!='REVIEW_MANIFEST.json'}}
    S.write(RUN/'REVIEW_MANIFEST.json',manifest)
    assert all(S.sha(p)==h for section in ('product','tests','evidence') for p,h in manifest[section].items())
    print('FINAL_SEAL_PASS',len(manifest['evidence']),'PROTECTED',len(protected),'FINGERPRINT',S.digest(product),
          'MANIFEST',S.sha(RUN/'REVIEW_MANIFEST.json'),flush=True)
