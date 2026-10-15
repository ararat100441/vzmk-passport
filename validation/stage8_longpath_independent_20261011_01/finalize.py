"""Write the final report, then seal and independently read back every entry."""
import json,hashlib,sys
from pathlib import Path
from datetime import datetime,timezone,timedelta
RUN=Path(__file__).resolve().parent;ROOT=RUN.parents[1]
sys.path.insert(0,str(RUN));from integrity import read,sha,save,hashes
SHORT=ROOT/'validation/stage8_li_261011_01'
q=read(RUN/'CANONICAL_RESULTS.json');assert q['PASS']==89 and q['FAIL']==0
integrity=read(RUN/'SOURCE_INTEGRITY.json');assert integrity['product_before_equals_after'] and not integrity['protected_data_changed'] and not integrity['read_errors']
product=read(RUN/'PRODUCT_AFTER.json');assert product==read(RUN/'PRODUCT_BEFORE.json')
assert all(sha(Path(p))==h for p,h in product.items())
excel=read(RUN/'FINAL_EXCEL_OWNERSHIP.json');assert not excel['live_owned']
matrix=read(RUN/'PREVIOUS_15_INDEPENDENT.json')['results'];raw=read(RUN/'RAW_JOURNAL_CHAIN.json')
left=read(RUN/'SHORT_WORKSPACE_LEFTOVERS.json')['retained_crash_or_adversarial_evidence']
date=datetime.now(timezone(timedelta(hours=4))).isoformat()
def link(name):return '['+name+'](<'+(RUN/name).as_posix()+'>)'
lines=[
'# ВЗМК — Этап 8 — FINAL INDEPENDENT REVIEW — LONG PATH FIX',
'',f'Завершено: {date}, Europe/Samara. Новый независимый пакет, начат 11.10.2026. Product и реальные источники READ-ONLY; все мутации XLSM — на новых искусственных TEST-копиях.',
'','**FINAL VERDICT: PASS. BLOCKER 0 / MAJOR 0 / MINOR 0. READY FOR USER E2E: YES.**',
'','## 1. ACTUAL PRODUCT SHA','',
'SHA-256 fingerprint карты 14 файлов: `'+q['product_fingerprint']+'`.',
'Это SHA canonical JSON полной карты абсолютных путей и SHA файлов (UTF-8, sort_keys, separators=comma/colon), не Git commit.',
'Карты до и после: '+link('PRODUCT_BEFORE.json')+', '+link('PRODUCT_AFTER.json')+'. Все 14 SHA неизменны.',
'','| Product file | Actual SHA-256 |','|---|---|',
]
for p,h in product.items():lines.append('| '+Path(p).name+' | `'+h+'` |')
lines += [
'', 'Независимо восстановленный полный diff побайтно по тексту совпал с предоставленным: '+link('INDEPENDENT_PRODUCT.diff')+', '+link('EXACT_DIFF_INDEPENDENT.json')+'. Изменены только core.py, native.py, recovery.py; добавлен shortcopy.py. storage.py, office.py, GUI, CLI и инструкция не менялись.',
'Авторские 80 PASS связаны с текущими SHA и 6 565 проверенными записями manifest: '+link('AUTHOR_80_BINDING.json')+', '+link('PRODUCT_BINDING.json')+'. Это отдельная авторская матрица. Независимая окончательная матрица — **89 PASS / 0 FAIL**: long 27 + short 27 + edges_final 34 + видимый GUI 1.',
'','## 2. INDEPENDENT LONG PATH TESTS','',
'Проверены реальные native Open источников длиной 259, 260, 261 и 370 символов. Основной длинный restore verified.xlsm имеет длину '+str(raw['long_restore_stage_lengths'])+', короткий — '+str(raw['short_restore_stage_lengths'])+'. Все PASS. Это проверка установленного Excel, не обещание универсального лимита Windows.',
'Защищённая копия побайтно совпадает с источником; SHA и volume_serial/file_index связаны в verified_copy. Отдельные внешние процессы не смогли write, обычный replace или native POSIX replace ни source, ни copy; rename временного parent также запрещён.',
'Protected DACL создаётся сразу и допускает только текущего пользователя и SYSTEM. Уникальная папка и exclusive xb исключают переиспользование; искусственная UUID-коллизия дала STOP до INTENT. Проверены реальные ACL-deny и отсутствующая, относительная, слишком длинная base и base-файл.',
'Проверки SHA/identity до native и после настоящего native Excel выполнены. Подмена байтов до native не дошла до Excel. Same-bytes identity-подмена вызвала STOP cleanup с сохранением чужой identity. После native отдельно инъецированы наблюдения изменённых SHA/identity source и copy: каждый раз STOP без возвращённого SUCCESS. Физическая внешняя запись при удерживаемых handles отдельно проверена реальными процессами и запрещена.',
'Обычное исключение сохраняет первичную причину и удаляет свои файл/каталог; foreign collision сохраняется и вызывает STOP. Preflight close — до INTENT/backup/изменения XLSM, restore и resume — до любых target rename. Нативный отказ preflight также оставил журнал и XLSM неизменными.',
'Evidence: '+link('edges_final/RESULTS.json')+', '+link('long/RESULTS.json')+', '+link('short/RESULTS.json')+'.',
'','## 3. ALL PREVIOUS 15 FAIL RESULTS','',
'Все 15 имён взяты из прежнего независимого RESULTS.json и повторены заново на двух длинах. Незавершённый старый manifest не использован как доказательство PASS. Полная связь case/SHA: '+link('PREVIOUS_15_INDEPENDENT.json')+'.',
'','| Сценарий | Ранее | Short | Long |','|---|---|---|---|',
]
for row in matrix:lines.append('| '+row['name']+' | FAIL | PASS | PASS |')
lines += [
'','## 4. BACKUP / RECOVERY VALIDATION','',
'Реальная смерть TEST Python (os._exit 73) в during_publish оставила повреждённый XLSM. Обычные recover и close обнаружили неоднозначность и отказали без replay. Проверены backup SHA/metadata и отказ для повреждённого/подменённого backup, повреждённого журнала и stage, неизвестной identity и валидного постороннего содержимого.',
'Пустое подтверждение не меняет данные. Явный operator restore возвращает точные BEFORE bytes, новую разрешённую identity, RECOVERED_ABORTED, decrements=0, closed=False. Повреждённые байты остаются в quarantine; прежние receipts/journal entries не перезаписаны. У восстановленной операции нет COMMITTED.',
'Шесть аварийных точек restore (до подготовки, chunk-copy, после intent, после quarantine, после publish, перед terminal) проверены в новых процессах с fresh preview и явным resume. Сохраняются более поздние COMMITTED; rollback уже завершённого закрытия запрещён. Полный AFTER после аварии финализируется без повторной записи и без rollback.',
'Независимо пересчитаны raw journal seals/previous links и связи restoration identity: '+link('RAW_JOURNAL_CHAIN.json')+'. Два намеренно испорченных журнала отдельно отвергнуты как ожидаемые отрицательные fixtures.',
'','## 5. XLSM PRESERVATION','',
'Самостоятельный Microsoft Excel 16 COM: normal Open с CorruptLoad=0, ReadOnly=True, FileFormat=52, AutomationSecurity=3, UpdateLinks=0. Источники не сохранялись. Восстановление — все ZIP parts побайтно равны backup, все 11 VBA streams равны.',
'До/после нового close отдельно прочитаны native Excel стили, font Bold/Size=13, number format=0, ColumnWidth=55, формулы, листы/видимость, Names, LinkSources и Connections. Стиль и структура неизменны; veryHidden сохранён. Формула SUM меняется по рассчитанному значению 30→23, выражение неизменно; =2+3 остаётся 5. Остатки 12/10→9/6.',
'Полное сравнение package разрешает только C24/C25 и ранее явным образом разрешённые save metadata/calculation caches; стили, отношения, структура и все 11 исходных/compiled VBA streams сохранены. Evidence: '+link('edges_final/native_styles_links_visibility_and_11_streams.json')+', '+link('NATIVE_PRESERVATION.json')+', '+link('long/NATIVE_EXACT.json')+'.',
'','## 6. NO DOUBLE DECREMENT','',
'Fresh process после restore даёт NO_PENDING_OPERATION. Старый pre-restore close-preview блокируется по FILE_ID_CHANGED. Новый preview и новое подтверждение закрывают один раз: 12−3=9, 10−4=6, ровно один COMMITTED. Повтор подтверждённого close-preview, новый процесс и переименованный request дают ALREADY_CLOSED_TEST без записи. Два одновременных close дают ровно один debit.',
'Регрессии 4/9: конкурентный внешний publish запрещён удерживаемым guard; если внешний writer успел до acquisition, его 4 сохраняется, старый расчёт 9 отклоняется до INTENT. Existing writer и writable mapping блокируют close. DEV-паспорт 9/2026 по-прежнему BLOCKED; official_issue=false.',
'','## 7. EXCEL COM CLEANUP','',
f'Проверены {excel["ownership_records"]} START/END ownership records с PID/creation_time и expected TEST workbook. Во всех normal/exception сессиях присутствуют независимые попытки CutCopyMode, Workbooks.Close и Quit. Включены инъекции первого cleanup-error и Workbooks.Close-error, первичный Workbook.Close/Save HRESULT, exact 0x80010108 и настоящий disconnect собственного Excel через Quit перед Save. Ошибки не маскируются SUCCESS.',
'Собственных живых Excel после освобождения COM proxy: **0**. Исходный неатрибутированный Excel PID 51208 с той же creation time сохранён. Глобального taskkill, принудительного завершения PID и очистки процессов во время тестов не было. Evidence: '+link('EXCEL_COM_RECORDS.json')+', '+link('FINAL_EXCEL_OWNERSHIP.json')+'.',
'Ошибка автора изучена независимо: cleanup завершил его собственный активный TEST Excel PID 70112, относящийся к refuse_valid_unknown_inplace в short_02. Ownership действительно подтверждён, но первоначальное объяснение orphan было неверным. Получен 0x800706BE STOP; backup цел, COMMITTED отсутствует. Это ошибка orchestration предварительного стенда; short_02 исключён. Тот же сценарий в наших short/long PASS. '+link('AUTHOR_CLEANUP_INCIDENT_INDEPENDENT.json')+'.',
f'Normal temporary copy и каталог удаляются (проверено тестами). После намеренного os._exit finally технически не выполняется; {len(left)} каталогов controlled crash/adversarial evidence оставлены под уникальными именами, не переиспользуются. В том числе намеренно созданные foreign identity/collision bytes сохранены. Неизвестные leftovers не удаляются рекурсивно. '+link('SHORT_WORKSPACE_LEFTOVERS.json')+'.',
'','## 8. GUI / CLI STATUS','',
'CLI PASS: реальные отдельные процессы restore-preview, cancel, confirmed restore и повторный запуск на коротких/длинных путях. Видимый GUI PASS через computer-use @oai/sky: настоящий chooser, emergency preview на verified path 260, default NO, Return cancel без SHA/journal изменений, Yes restore с native Excel, RECOVERED_ABORTED/0, новый preview 12−3=9 и 10−4=6, confirmation=false. Messagebox/handlers не патчились. Собственное окно и Python завершены.',
'Evidence: '+link('GUI_VALIDATION.json')+', '+link('GUI_01_PENDING.png')+', '+link('GUI_02_DEFAULT_NO.png')+', '+link('GUI_04_RECOVERED.png')+', '+link('GUI_05_FRESH_PREVIEW.png')+'. Ручной User E2E самого пользователя остаётся отдельным следующим действием.',
'','## 9. SOURCE INTEGRITY','',
f'Повторно проверены {integrity["checked"]} исходных путей, включая product, реальные книги/реестры/документы/ВОМ/справочники/шаблоны и прежние evidence. Protected data changed: 0; read errors: 0; все 14 product SHA до/после равны.',
'Технические изменяемые index heartbeat, если присутствуют: '+str(len(integrity['volatile_index_heartbeat']))+'. Они перечислены отдельно и не относятся к бизнес-данным. '+link('SOURCE_INTEGRITY.json')+', '+link('PROTECTED_BEFORE.json')+', '+link('PROTECTED_AFTER.json')+'. Авторские 28 911 protected entries также полностью совпали с текущими до тестов: '+link('HISTORICAL_CLAIM.json')+'.',
'Новые записи находятся только в данном evidence/TEST пакете, новом short TEST root и уникальных защищённых short-workspaces. Исходные Stage 1–7, реальные источники и старые receipts не изменялись.',
'','## 10. BLOCKER / MAJOR / MINOR','',
'**BLOCKER 0 / MAJOR 0 / MINOR 0.** Последний MAJOR на verified.xlsm длиной 260 устранён и независимо воспроизведён как PASS.',
'Предварительный независимый edges: 31 PASS / 3 FAIL были ошибками review harness (ожидаемый label безопасного STOP, невалидная SDDL-маска теста, нерешённый виртуализированный LOCALAPPDATA в прямом COM Open). Исходный прогон сохранён, изменения только тестового скрипта связаны SHA; полностью повторён edges_final: 34 PASS / 0 FAIL. '+link('EDGES_HARNESS_REPAIR.json')+'. Product не исправлялся.',
'Первый проход сборщика audit остановился при чтении identity каталога: тестовый helper вызвал file_identity без FILE_FLAG_BACKUP_SEMANTICS. Исправлен только metadata-reader сборщика; исходный traceback сохранён в '+link('AUDIT_01.stderr.txt')+'. Повторный raw audit завершился PASS; функциональные тесты и product не изменялись.',
'','## 11. NOT_RUN','',
'Ручной User E2E пользователем — NOT_RUN; готов к отдельному выполнению. Официальное закрытие DEV 9/2026 и реальные списания запрещены и не выполнялись. Спонтанная историческая причина RPC_E_DISCONNECTED остаётся UNKNOWN; её не заменяет успешная fault injection. Длины вне проверенной среды и произвольные сетевые файловые системы не заявляются проверенными.',
'Семантический index вернул INTERNAL_ERROR; для узкого diff использованы точные source reads и отдельная проверка snapshot/SHA. Это ограничение инструмента поиска не прервало независимую проверку.',
'','## 12. FINAL VERDICT','',
'**PASS — 89 PASS / 0 FAIL. BLOCKER 0 / MAJOR 0 / MINOR 0.** Полная матрица: '+link('CANONICAL_RESULTS.json')+'.',
'Новый REVIEW_MANIFEST.json создаётся последним после завершения workers/GUI/COM cleanup, включает product и все новые evidence/TEST files. MANIFEST_VERIFICATION.json независимо перечитывает все entries и связывает SHA самого manifest; он намеренно не включает себя. Старый незавершённый manifest не является основанием этого PASS.',
'','## 13. READY FOR USER E2E','',
'**YES**, для отдельного ручного DEV/TEST User E2E. Это не официальное закрытие 9/2026 и не разрешение Этапа 9.',
'','STOP. Product code не изменён. Этап 9, EXE и deployment не выполнялись.','',
]
(RUN/'FINAL_REPORT.md').write_text('\n'.join(lines),encoding='utf8',newline='')
paths={str(p.resolve()) for base in (RUN,SHORT) for p in base.rglob('*') if p.is_file() and '__pycache__' not in p.parts and p.name not in ('REVIEW_MANIFEST.json','MANIFEST_VERIFICATION.json')}
for folder in left:paths.update(str(Path(p).resolve()) for p in folder['files'])
evidence=hashes(paths);assert all(isinstance(h,str) for h in evidence.values())
m={'schema':'STAGE8_LONGPATH_INDEPENDENT_COMPLETED_V1','completed_at':date,'product':product,'product_fingerprint':q['product_fingerprint'],'evidence':evidence,'canonical_PASS':89,'canonical_FAIL':0,'previous_15_short_PASS':15,'previous_15_long_PASS':15,'protected_checked':integrity['checked'],'protected_data_changed':0,'live_owned_Excel':0,'independent_review_complete':True,'final_verdict':'PASS','ready_for_user_E2E':True}
save('REVIEW_MANIFEST.json',m)
errors=[]
for group in ('product','evidence'):
    for p,h in m[group].items():
        if sha(Path(p))!=h:errors.append(p)
assert not errors,errors
save('MANIFEST_VERIFICATION.json',{'status':'PASS','manifest_sha256':sha(RUN/'REVIEW_MANIFEST.json'),'checked_product_entries':len(product),'checked_evidence_entries':len(evidence),'errors':errors,'final_report_sha256':sha(RUN/'FINAL_REPORT.md'),'complete':True,'self_excluded_by_design':True})
print('SEALED_PASS',len(evidence),'manifest',sha(RUN/'REVIEW_MANIFEST.json'),flush=True)
