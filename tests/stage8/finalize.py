"""Seal Stage 8 evidence only after targeted tests and protected-source audit."""
import json
from pathlib import Path
import sys
from stage8_closure import core,storage as S
from baseline import snapshot

run=Path(sys.argv[1]).resolve()
origin=Path(sys.argv[2]).resolve()
before=S.read(origin/'SOURCE_BEFORE.json');after=snapshot()
changes=[{'path':p,'before':before.get(p),'after':after.get(p)} for p in sorted(set(before)|set(after)) if before.get(p)!=after.get(p)]
S.write(run/'SOURCE_AFTER.json',after)
S.write(run/'SOURCE_INTEGRITY.json',{'status':'PASS' if not changes else 'FAIL','guarded_before':len(before),
    'guarded_after':len(after),'changes':changes,'before_path':str(origin/'SOURCE_BEFORE.json'),
    'before_sha256':S.sha(origin/'SOURCE_BEFORE.json'),'after_path':str(run/'SOURCE_AFTER.json'),
    'after_sha256':S.sha(run/'SOURCE_AFTER.json'),
    'scope':['Pasport including preexisting validation/DEV/receipts','Документы/Гараж казахстан',
             'Работа/Другие шаблоны','Работа/ВЗМК.xlsx','Работа/УЗК.xlsm'],
    'exclusions':['only newly authorized Stage 8 code/tests/docs/evidence','__pycache__','.codebase-index']})
main=S.read(run/'targeted/RESULTS.json');extra=S.read(run/'additional/RESULTS.json');stable=S.read(run/'OFFICE_STABILITY.json')
product=core.code_bindings()
code_equal=(product==S.read(run/'targeted/PRODUCT_BEFORE.json')==S.read(run/'targeted/PRODUCT_AFTER.json')
            ==extra['product_before']==extra['product_after']==stable['product'])
assert not changes,changes[:10]
assert code_equal and main['failed']==0 and extra['failed']==0 and stable['status']=='PASS'
native=S.read(run/'targeted/POSITIVE_NATIVE.json')
cases=main['cases']+extra['cases']
required=['successful_close_backup_debit_repeat_xlsm','insufficient_balance','mandatory_document_missing',
    'no_uzk_required_no_registry_write','input_changed_after_preview','failure_before_save',
    'failure_at_native_save_io_boundary','torn_save_result_rejected','failure_after_save_before_publish',
    'process_death_after_publish_fresh_recovery','two_simultaneous_processes',
    'passport_9_2026_and_official_promotion_forbidden']
assert set(required).issubset({r['name'] for r in cases if r['status']=='PASS'})
newfiles=[*sorted((S.ROOT/'stage8_closure').glob('*.*')),S.ROOT/'stage8.py',S.ROOT/'STAGE8_DEV.md',
          S.ROOT/'Запустить Этап 8 DEV.cmd',*sorted((S.ROOT/'tests/stage8').glob('*.py'))]
newfiles=[p for p in newfiles if p.is_file()]
code_inventory={str(p):S.sha(p) for p in newfiles}
S.write(run/'NEW_FILES.json',code_inventory)
S.write(run/'TARGETED_SUMMARY.json',{'status':'PASS','passed':len(cases),'failed':0,'cases':cases,
    'office_stability_rounds':len(stable['rounds']),'product_bindings_equal':code_equal,
    'official_close':'FORBIDDEN_NOT_RUN','real_uzk_debit':'FORBIDDEN_NOT_RUN',
    'manual_user_e2e':'NOT_RUN','full_stages_1_7_regression':'FORBIDDEN_NOT_RUN',
    'pdf_exe_deploy_stage9':'FORBIDDEN_NOT_RUN'})

table='\n'.join('| '+r['name']+' | '+r['status']+' |' for r in cases)
filelist='\n'.join('- `'+str(p.relative_to(S.ROOT))+'`' for p in newfiles)
report=f'''# ВЗМК — Этап 8 — отчёт для независимой проверки Work 08

Дата: 10.10.2026 (Europe/Samara). READY FOR WORK 08 REVIEW: **YES**.
Объём готовности: согласованный DEV/TEST. Официальное закрытие и реальные
списания запрещены и не выполнялись. После передачи отчёта STOP; Этап 9 не начат.

## 1. Что реализовано

Отдельный локальный модуль закрытия, предварительная проверка, подтверждение
пользователя, проверяемое резервирование, native-редактирование XLSM, однократное
списание, устойчивый журнал, обнаружение незавершённых операций и восстановление.
Есть самостоятельные Tk DEV-интерфейс, CLI и генератор изолированного стенда.
AI/интернет/API в работе программы не нужны. Принятые этапы не переписаны.

## 2. Как работает закрытие

DEV-интерфейс показывает проект, номер, отгрузку, документы, обязательный состав,
статус проверки, подтверждённое количество УЗК и остатки по маркам до/после.
Требуются checkbox и отдельное окно подтверждения с отменой по умолчанию.
CLI требует номер паспорта и печать конкретного preview. Повторно проверяются
SHA всех входов и Stage 8 кода, принадлежность документов и контроль.
Ошибки, TEST/SAMPLE_ONLY вместо подтверждения или отсутствие документов дают STOP.

В этой поставке состояние успешного закрытия только CLOSED_TEST,
official_issue=false. Настоящие документы 9/2026 не повышаются до OFFICIAL.
Положительные тесты используют вымышленные документы JSON и подтверждения,
созданные с нуля, а не подменённые receipts этапов 5–7.

## 3. Как работает списание УЗК

Реально редактируется только тестовая XLSM внутри стенда. Списывается точное
подтверждённое количество изделий по выбранным маркам. Положительный сценарий:
S8-A 12−3=9, S8-B 10−4=6; всего 7 изделий. Вся отгрузка или длина швов не списываются.
Ноль допустим, отрицательное значение запрещено. Пустое исходное C означает 1
по принятому контракту Этапа 5; дубли выбранной марки дают STOP.
Ключ закрытия — нормализованный проект + паспорт, независимо от пути запроса
и отгрузки. Повторные вызовы/новый preview/новый процесс дают ALREADY_CLOSED_TEST
без записи. При явно отсутствующей необходимости УЗК SHA XLSM неизменен,
backup/save/replace XLSM отсутствуют. Реестр закрытий — новый отдельный журнал;
реестр ВЗМК не получает новых колонок или записей.

## 4. Резервные копии

Перед подготовкой изменяемой книги создаётся `backups/<operation_id>/УЗК_TEST.xlsm`.
Копия сбрасывается на диск, её SHA-256 сравнивается с защищённым оригиналом.
Отдельный SHA256.json и его SHA связаны с журналом. Ошибка записи или несовпадение
SHA запрещают сохранение/публикацию. Предыдущие копии сохраняются. Проверены два
последовательных паспорта и две независимые копии; первая осталась неизменной.
Положительная копия отдельно открыта Microsoft Excel read-only и совпала с SHA «до».

## 5. Сбои и восстановление

Состояния: INTENT → BACKUP_VERIFIED → STAGED → APPLYING → COMMITTED; при отмене
восстановления — ABORTED. События только добавляются, имеют последовательность,
SHA-печать и SHA предыдущего файла, записываются с flush/fsync перед публикацией.
Windows lock сериализует операции. Входы, код и книги защищены deny-write handles;
на атомарном переходе защита удерживается и на новом файле. Публикация —
SetFileInformationByHandle / FileRenameInfoEx, same-volume, без fallback delete/copy;
DACL переносится. При неподдерживаемом native-вызове — STOP.

Восстановление не сохраняет и не заменяет XLSM. Если SHA равен «до», операция
отменяется без списания. Если APPLYING и SHA точно равен проверенному «после»,
после проверки копии/evidence/native Excel только дописывается COMMITTED.
Неоднозначность, повреждение копии/журнала/evidence или потеря журнала при изменённой
книге дают STOP без повтора. Проверены реальные os._exit(73) до сохранения и после
публикации и восстановление в свежем процессе; повторного списания нет.

## 6. Созданные файлы

Все перечисленные файлы новые; ранее принятый код не изменён.

{filelist}

`stage8_closure/office.py` владеет только своими Excel.Application. COM apartment
живёт весь процесс; каждый экземпляр Excel явно закрывается. Исходный helper
Этапа 1 не изменялся. Probe-скрипты сохранены как инженерная диагностика.

## 7. Положительные проверки

- Успешное закрытие, точные 7 изделий, сохранение макросов/формул/стилей/структуры.
- Проверяемая копия, её независимое открытие Excel, сохранение старых копий.
- Три повторных вызова и повторное открытие без записи и списания.
- Закрытие без требуемого УЗК без изменения XLSM; списание до нулевого остатка.
- Два различных паспорта с сохранением обеих копий и правильными остатками.
- Восстановление в новом процессе после реальной смерти процесса.
- GUI: полное preview, checkbox, modal cancel/accept, блокировка повторного действия.
- Дополнительная стабильность Office: {len(stable['rounds'])} полных циклов создания/закрытия
  подряд в одном Python-процессе, каждый с отдельными экземплярами Excel.

## 8. Отрицательные проверки

Недостаток остатка; отсутствующий ВИК; неподтверждённый контроль; SAMPLE_ONLY;
неполное количество; документ другой отгрузки; изменённый вход или XLSM;
подмена preview; невозможность/повреждение backup; ошибка Save и повреждённый
результат Save; сбой до публикации; две одновременные попытки; внешний writer;
дубли строки; повреждённый/потерянный журнал; неизвестное состояние восстановления;
изменённый preservation evidence; внешние пути; официальный режим и 9/2026 — PASS
как ожидаемый безопасный отказ без неконтролируемой записи.

Полная матрица targeted-тестов: **{len(cases)}/{len(cases)} PASS**, FAIL=0.

| Проверка | Результат |
|---|---|
{table}

## 9. Сохранность XLSM через Microsoft Excel

Excel {native['after']['excel_version']}; FileFormat=52; HasVBProject=true;
нормальное открытие CorruptLoad=0, ReadOnly=true, VBA/события/links отключены.
Все {len(native['after']['vba_streams'])} VBA-потоков, включая исходный и скомпилированный
код, побайтно совпадают. OLE-контейнер Excel может переставить секторы, поэтому
его общий binary SHA меняется; это явно отражено в preservation evidence.
Никакой ZIP-перезаписи VBA или XLSM через openpyxl не выполнялось.

Сравниваются все части OOXML-пакета: значения вне списания, формулы, styles,
листы/visibility, связи, имена, размеры/настройки и прочая структура. Разрешены
только конкретные C-значения и перечисленные метаданные native Save (время/автор
изменения, calcPr/calcChain, каталог сохранения, coauthor documentId, пустая запись
built-in functionGroups=19, вычисленные caches без изменения формул).
Две специально добавленные в fixture формулы сохранены; зависимая сумма
пересчиталась с 30 до 23, независимая формула осталась 5. veryHidden лист и стили
сохранились. Независимое native read-only открытие backup тоже PASS.

## 10. Сохранность исходных файлов

BEFORE={len(before)}, AFTER={len(after)}, changes={len(changes)}. SHA-256 всех защищённых
файлов совпали. Включены весь исходный Pasport (код этапов 1–7, DEV, справочники,
шаблоны внутри проекта, прежние validation и receipts), документы Гараж Казахстан
с исходными ВОМ/реальными паспортами, Другие шаблоны и два рабочих реестра ВЗМК/УЗК.
Исключены только новые разрешённые файлы Stage 8, технический __pycache__ и индекс.
Побайтные таблицы и границы доказательства — SOURCE_BEFORE/AFTER/INTEGRITY.json.
Нет необъяснённых изменений источников. Тестируемые Stage 8 code bindings
совпадают до и после обеих серий и Office stability; полный список NEW_FILES.json.

## 11. Проблемы и ограничения

В ходе разработки в новых Stage 8 модулях устранены локализация числового формата,
жёсткая привязка к первому листу, ложное отличие OLE-контейнера при одинаковых VBA,
конфликт блокировок с path-based replace и повторный CoUninitialize при живых COM
proxy. Два первоначальных диагностических запуска остановились; изменённый реестр
не публиковался, следы сохранены в stage8_20261010_01/02. Оставшиеся после них только
два созданных нами automation Excel-процесса были адресно закрыты.

В STDERR основного успешного запуска диагностический faulthandler зарегистрировал
Windows HRESULT 0x80010108 во время одного Excel Save. Python завершился кодом 0,
все 29 сценариев, native checks и последующие 7 дополнительных проверок прошли.
Это диагностический факт; он не скрыт и не засчитан как отдельная проверка.
Отдельный OFFICE_STABILITY.json и его STDOUT/STDERR сохраняют последующую проверку.

Нет подтверждённых реальных контрольных данных: официальный режим намеренно
заблокирован; позитивные документы искусственные, не пригодны для выдачи.
Тестируемый runtime — локальный Windows/NTFS + установленный Excel + Python 3.12.
Физическое отключение питания/неисправность диска и ручной пользовательский E2E
NOT_RUN; моделирование I/O-ошибок и torn Save не является аппаратным испытанием.
Полная регрессия этапов 1–7, PDF, EXE, deployment, Stage 9 — NOT_RUN по заданию.
SHA-цепь — локальная целостность, не электронная подпись и не защита от злонамеренного
пересчёта всех файлов администратором. Автоматического разрушительного rollback нет.

## 12. Точные пути результатов

- Канонический пакет: `{run}`.
- Исходный BEFORE: `{origin/'SOURCE_BEFORE.json'}`.
- Матрица: `TARGETED_SUMMARY.json`; подробные 29 протоколов: `targeted/`.
- Семь дополнительных протоколов и GUI_TEXT.json: `additional/`.
- Native XLSM/backup evidence: `targeted/POSITIVE_NATIVE.json`.
- Повторные native циклы: `OFFICE_STABILITY.json`, `STABILITY_STDOUT/STDERR.txt`.
- Стенды: `sandboxes/<uuid>/`; точный UUID каждого случая записан в его протоколе.
- Копии, журнал, preservation: внутри соответствующего стенда в backups/journal/transactions.
- Контроль исходников: `SOURCE_AFTER.json`, `SOURCE_INTEGRITY.json`.
- Заключительный seal всего пакета и новых файлов: `REVIEW_MANIFEST.json` (пишется последним).

## 13. Решение и граница

READY FOR WORK 08 REVIEW: **YES** — только в разрешённом DEV/TEST объёме.
Реальное официальное закрытие: FORBIDDEN / NOT_RUN. Реальные списания: 0.
Этапы 1–7 неизменны. Этап 9 не начат. **STOP** после передачи независимой проверке.
'''
(run/'FINAL_REPORT.md').write_text(report,encoding='utf8')
evidence={str(p):S.sha(p) for p in sorted(run.rglob('*')) if p.is_file() and p.name!='REVIEW_MANIFEST.json'}
manifest={'schema':'STAGE8_REVIEW_MANIFEST_V1','created_at':S.now(),'ready_for_work_08_review':'YES',
    'scope':'ISOLATED_DEV_TEST_ONLY','product':code_inventory,'evidence':evidence,
    'source_before':{'path':str(origin/'SOURCE_BEFORE.json'),'sha256':S.sha(origin/'SOURCE_BEFORE.json')},
    'source_integrity':{'guarded':len(before),'changes':[]},'targeted':{'passed':len(cases),'failed':0},
    'office_stability_rounds':len(stable['rounds']),'official_issue':False,'real_registry_debits':0,'stage9':'NOT_STARTED',
    'external_documentation':[
        'https://learn.microsoft.com/en-us/windows-hardware/drivers/ddi/ntifs/ns-ntifs-_file_rename_information',
        'https://learn.microsoft.com/en-us/windows/win32/api/fileapi/nf-fileapi-setfileinformationbyhandle']}
S.write(run/'REVIEW_MANIFEST.json',manifest)
# Read-only post-seal verification. No further files are written into this package.
for bindings in (code_inventory,evidence,{str(origin/'SOURCE_BEFORE.json'):S.sha(origin/'SOURCE_BEFORE.json')}):
    for path,expected in bindings.items():
        assert S.sha(path)==expected,path
print(json.dumps({'ready':'YES','passed':len(cases),'failed':0,'guarded_sources':len(before),
    'source_changes':len(changes),'new_files':len(newfiles),'evidence_files':len(evidence),
    'manifest_sha256':S.sha(run/'REVIEW_MANIFEST.json'),'report':str(run/'FINAL_REPORT.md')},ensure_ascii=False),flush=True)
