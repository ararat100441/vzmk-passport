"""Finalize R2-only evidence. REVIEW_MANIFEST.json is written last."""
from common import *
import ast,datetime
b=load(E/'baseline.json');before=load(E/'final_tests_product_BEFORE.json');after=load(E/'final_tests_product_AFTER.json');current={p:sha(p) for p in before}
assert before==after==current
all_current={p:sha(p) for p in b['guarded']};changed={p:{'reviewed_baseline':h,'current':all_current[p]} for p,h in b['guarded'].items() if all_current[p]!=h}
assert set(changed)=={str(R/'akz_stage6/core.py')}
tests=load(E/'targeted_tests.json');assert len(tests['results'])==27 and all(t['status']=='PASS' for t in tests['results'])
assert tests['audit']=={'production_attempts':[],'network_attempts':[]}
returned=load(E/'native_GUI_after_return_lock_still_held.json');retry=load(E/'native_GUI_release_cleanup_retry.json');modal=load(E/'native_GUI_message_while_locked.json')
assert returned['handle_live_after_GUI_return'] and returned['no_partial_published'] and returned['old_good_unchanged']
assert retry['pending_removed_by_product'] and retry['ownership_records_removed'] and retry['unrelated_unchanged'] and retry['old_good_unchanged']
assert modal['handle_live_during_modal'] and modal['message'].startswith('FILE_LOCKED') and 'Закройте файл в Excel' in modal['message'] and 'OTHER_ERROR' not in modal['message']
assert 'FILE_LOCKED' in (E/'native_GUI_lock.txt').read_text('utf-8')
assert not list(E.rglob('.pending-*')) and not list(E.rglob('.stage6-pending-*.json'))
assert all(sha(p)==h for p,h in retry['sha256'].items())
old_pair=R/'validation/stage6_fix_r1_r2_20261009_094707/NEW_TEST_9_2026'
old_expected={ACTNAME:'b640796570a8d04b6bca404205d99eae4c7005c700587812ed25bc99e715eb1a',APPNAME:'7766c92df82df2484eb82cd6917769f5d9ede912b3f0e8f26e3f8aaabdd357f2'}
assert all(sha(old_pair/n)==h for n,h in old_expected.items())
def defs(t):return {n.name:ast.dump(n,include_attributes=False) for n in ast.parse(t).body if isinstance(n,(ast.FunctionDef,ast.ClassDef))}
x=defs((E/'BEFORE_core.py').read_text('utf-8'));y=defs((R/'akz_stage6/core.py').read_text('utf-8'))
changed_definitions=[n for n in x if x[n]!=y[n]];assert changed_definitions==['FileOperationError','generate']
save('semantic_scope.json',{'changed_existing_definitions':changed_definitions,'unchanged_existing_definitions':[n for n in x if x[n]==y[n]],'R1_write_application_unchanged_AST':True,'calculation_template_validation_unchanged_AST':True,'GUI_unchanged':True})
save('tested_product_hashes.json',{'BEFORE':before,'AFTER':after,'CURRENT':current,'BEFORE_EQUALS_AFTER_EQUALS_CURRENT':True,'meaning':'Final tests snapshots, all after the authorized R2-only source change; reviewed pre-fix baseline is separate in baseline.json'})
save('source_old_artifact_safety.json',{'reviewed_guarded_count':len(all_current),'unchanged_count':len(all_current)-1,'intentional_changes':changed,'CURRENT':all_current,'old_pair':{str(old_pair/n):sha(old_pair/n) for n in old_expected}})
save('pending_ownership_assessment.json',{'persisted_schema':'STAGE6_OWNED_PENDING_V1','ownership_fields':['UUID','exact output root','exact staging path','directory dev/inode','registered file dev/inode'],'reparse_files_and_directories_rejected':True,'registered_filenames':[ACTNAME,APPNAME,'PAIR.json'],'unknown_pending_STOP_test':'PASS','unrelated_file_inside_owned_pending_STOP_test':'PASS','replaced_expected_filename_identity_STOP_test':'PASS','tampered_ledger_path_STOP_test':'PASS','fresh_process_recovery_after_release':'PASS','GUI_retry_product_cleanup_without_harness_cleanup':'PASS','recursive_deletion':False,'bounded_probe_no_wait_loop':True})
report=f'''1. ACTUAL STATE — R2 FIX COMPLETE; TARGETED VALIDATION PASS
R1 CLOSED/PASS сохранён. Исправлялся только S6-P1-R2.
Exact independent rereview: {b['review_manifest']}
До изменений все {len(all_current)} baseline fingerprints совпали, включая expected core/gui SHA.
Pre-fix core:8dffda645c6370e2e284beaad44354b361ef9bd680ddfe190a74b92a404ada10.
GUI неизменён:9a52745f321bc5c5036ae11c83becb6ceae74cf648dd126896e8030827772b1f.

2. ROOT CAUSE R2
До правки на отдельной копии повторён Work failure: здоровый APPLICATION под Windows CreateFile
GENERIC_READ,share=0 удерживался до фактического возврата Launcher.generate.
core.native_open выдавал FILE_LOCKED, но finally/shutil.rmtree выбрасывал PermissionError WinError32.
GUI показывал OTHER_ERROR, оставался locked pending. Evidence:baseline_failure_reproduced.json.

3. FIX R2
Cleanup больше не заменяет primary exception: повторно выбрасывается тот же объект primary;
код, путь, cause и понятный текст сохранены. Secondary exception хранится в primary.cleanup_error.
Для приложения GUI-текст: «Файл приложения АКЗ занят или заблокирован. Закройте файл в Excel и повторите операцию».
Новая собственная staging папка имеет UUID и durable sidecar .stage6-pending-UUID.json в output root.
Запись содержит schema,UUID,exact root/path,directory identity,NTFS identities каждого зарегистрированного файла.
После собственного Excel Save identity обновляется, поскольку Excel может заменить файл при сохранении.
Перед следующей генерацией recover_pending проверяет записи и принадлежность. Неизвестная папка,
постороннее содержимое,подмена identity/пути,reparse или некорректная запись дают PENDING_OWNERSHIP_STOP.
Сначала проверяются все файлы; при lock никакой файл staging не удаляется, возвращается FILE_LOCKED.
После release удаляются только зарегистрированные собственные файлы; rmdir удаляет только пустую папку.
Recursive delete отсутствует. Чужой процесс не завершается; блокировка не обходится; бесконечного ожидания нет.
Пара по-прежнему публикуется rename только после полной проверки ACT+APPLICATION и повторной проверки inputs.

4. CHANGED FILES
Единственный изменённый product file: {R/'akz_stage6/core.py'}
CURRENT SHA-256 {current[str(R/'akz_stage6/core.py')]}
GUI,write_application,R1,расчёты,шаблоны и оформление не менялись.
Diff:product_changes.diff. Новые fixtures/drivers/evidence только в этой папке;
служебный указатель:validation/STAGE6_R2_FINAL_CURRENT.txt.

5. LOCK HELD THROUGH GUI TEST — PASS
Реальная кнопка «Сформировать TEST-пару» нажата через native Sky UI.
CreateFile share=0 оставался открытым во время настоящего STOP/messagebox и после возврата GUI handler.
Пользовательский текст FILE_LOCKED,инструкция и полный путь визуально прочитаны; OTHER_ERROR отсутствует.
Evidence:native_GUI_lock_0.jpg/native_GUI_lock_1.jpg/native_GUI_lock.txt;
native_GUI_message_while_locked.json;native_GUI_after_return_lock_still_held.json.
Lock освобождён только после after-return evidence. Fake handler test отдельно подтверждает тот же сценарий.

6. PRIMARY ERROR PRESERVATION — PASS
Реальные primary FILE_LOCKED и secondary cleanup FILE_LOCKED сохранены отдельно.
Injected cleanup exception дополнительно проверена для FILE_LOCKED,ACCESS_DENIED,INVALID_XLSX,
OUTPUT_ALREADY_EXISTS_STOP,INPUT_CHANGED_AFTER_PREVIEW_STOP,OTHER_ERROR: объект primary не подменяется.
Сохранена исходная cause chain. Результаты в targeted_tests.json.

7. PENDING CLEANUP AFTER RELEASE — PASS
До release повторная операция возвращает FILE_LOCKED; file identities/size/mtime и ledger SHA неизменны.
После release принадлежность подтверждена, выполнена контролируемая очистка.
Durable recovery дополнительно выполнено новым Python процессом без in-memory registry.
Неизвестный pending,посторонний XLSX внутри собственного pending,замена ожидаемого файла новым file identity
и изменённый path ledger отдельно дают STOP до удаления. Пользовательские fixtures сохранены.
Evidence:pending_ownership_assessment.json и targeted_tests.json.

8. RETRY GENERATION — PASS
В настоящем GUI сценарии cleanup выполнялся самим продуктом при штатном generate после release,
без ручной очистки harness. Создана COMPLETE_TEST_PAIR:
{retry['complete_pair']}
{chr(10).join(f'{p}: {h}' for p,h in retry['sha256'].items())}
Все эти файлы — isolated validation outputs. Новая реальная DEV-пара не создавалась.
Старый успешный output/unrelated.xlsx не изменены. Stale pending и sidecar после retry отсутствуют.

9. ERROR CLASSIFICATION REGRESSION — PASS
FILE_LOCKED:реальный Windows32 при share=0; правильный GUI и pending retry.
ACCESS_DENIED:адресный Windows error5 unit,сообщение об отсутствии доступа.
INVALID_XLSX:реальные corrupt bytes/native Excel STOP,не классифицируется как lock,cleanup завершён.
OUTPUT_EXISTS:существующий target не перезаписан.
INPUT_CHANGED:каждый из5 isolated inputs после preview вызывает прежний STOP.
OTHER_ERROR:проверен valid XLSX+generic error; cleanup не маскирует остальные primary категории.

10. R1 SMOKE — PASS, NO R1 FIX
Native read-only saved isolated pair N3:4:6 visible,7:20 hidden,21 visible.
TOTAL3 /105.26kg /3.989m2; SUM C21/E21/G21 сохранены.
N17:полная согласованная pair,все17 рабочих строк видимы;N18:capacity STOP без усечения/output.
write_application AST идентичен reviewed baseline. Расчёты и оформление не переделывались.

11. SOURCE / OLD PAIR SAFETY — PASS
{len(all_current)-1}/{len(all_current)-1} guarded fingerprints неизменны; единственное разрешённое изменение — core.py.
Включены DEV registries/profiles,realcatalog,selected VOM,original АКЗ templates,Stage1–5,old outputs/evidence.
Existing reviewed pair неизменна:
{old_pair/ACTNAME}: {old_expected[ACTNAME]}
{old_pair/APPNAME}: {old_expected[APPNAME]}
Полные current source/old hashes:source_old_artifact_safety.json. Python forbidden open/socket attempts0;
web/network tools0. Production registries не использовались; системный packet capture не выполнялся.

12. TARGETED VALIDATION — PASS
27/27 адресных сценариев PASS + настоящий GUI persistent-lock/release/cleanup/retry PASS.
Прежний unit suite16/16 PASS. Все mutation tests только isolated fixtures.
Harness attempt1 после успешной проверки primary error попытался прочитать locked bytes для SHA,
что share=0 корректно запретил. Исправлен только harness: сравнение ID/size/mtime до release и SHA после.
Первый запуск сохранён,остаток после release очищен продуктом;полный suite повторён в sandbox_run2.
Evidence:HARNESS_RECOVERY.json,targeted_tests_attempt1.*,targeted_tests.json,unit_tests.txt.

13. EVIDENCE INTEGRITY — PASS
Final-tests product BEFORE=AFTER=CURRENT для всех7 product/launcher/docs/test файлов.
Pre-fix reviewed baseline сохранён отдельно; он намеренно отличается только новым core.py.
Baseline,точная before copy,minimal diff,real Windows lock,GUI,ownership,cleanup,retry,current hashes сохранены.
REVIEW_MANIFEST.json записан последним после этого report и проверен чтением/SHA всех included files.
Ни один прежний PASS не используется вместо новых persistent-lock проверок.

14. NOT_RUN
Независимый targeted Work rereview;full Stage1–5 regression;real ACL mutation;physical print;
новая реальная DEV-пара/новый native print preview;official issuance;production;Stage7;АКЗ1/2/3;
system packet capture. Index tools отсутствуют,использовались адресные чтения.

15. READY FOR TARGETED WORK REREVIEW — YES
Только S6-P1-R2 исправлен и адресно проверен. Независимый вердикт ещё не получен.
STOP. Stage7 и production не выполнялись.
'''
(E/'FIX_REPORT.txt').write_text(report,encoding='utf-8')
files={str(p):sha(p) for p in E.rglob('*') if p.is_file() and p.name!='REVIEW_MANIFEST.json'}
manifest={'kind':'R2-only implementation validation, not independent rereview','created_at':datetime.datetime.now().astimezone().isoformat(),'scope':['S6-P1-R2'],'R1':'CLOSED_UNCHANGED_SMOKE_PASS','reviewed_baseline_manifest':b['review_manifest'],'reviewed_baseline_manifest_sha256':sha(b['review_manifest']),'reviewed_product':b['product'],'final_tests_product':load(E/'tested_product_hashes.json'),'guarded_current':all_current,'changed_product_files':changed,'targeted_tests':27,'unit_tests':16,'native_GUI_persistent_lock_and_retry':'PASS','ready_for_targeted_Work_rereview':True,'files':files,'written_last':True}
save('REVIEW_MANIFEST.json',manifest)
assert all(sha(p)==h for p,h in load(E/'REVIEW_MANIFEST.json')['files'].items())
assert all(sha(p)==h for p,h in all_current.items())
print('FINAL PASS',len(files),'evidence files; manifest SHA256',sha(E/'REVIEW_MANIFEST.json'))
print('REPORT',str(E/'FIX_REPORT.txt'))
