from common_review import *
from datetime import datetime,timezone
b=load(E/'baseline_current.json');current={p:sha(p) for p in b['guarded']}
assert current==b['guarded'],'CURRENT CHANGED DURING REREVIEW STOP'
product={p:sha(p) for p in b['product']['CURRENT']}
assert product==b['product']['BEFORE']==b['product']['AFTER']==b['product']['CURRENT']
assert {n:sha(PAIR/n) for n in EXPECTED}==EXPECTED
t=load(E/'independent_tests.json');assert t['summary']=={'PASS':30,'FAIL':0} and t['units']=={'run':16,'pass':True}
assert not t['audit']['production_attempts'] and not t['audit']['network_attempts']
modal=load(E/'native_GUI_modal_lock_held.json');after=load(E/'native_GUI_after_return_lock_held.json');retry=load(E/'native_GUI_release_product_cleanup_retry.json')
assert modal['same_primary_object'] and modal['lock_live_during_actual_modal']
assert modal['proof'][0]['primary_code']=='FILE_LOCKED' and modal['proof'][0]['primary_cause_winerror']==32
assert modal['cleanup_diagnostic']['code']=='FILE_LOCKED' and modal['cleanup_diagnostic']['cause_winerror']==32
message=modal['messages'][0]['message'];assert message.startswith('FILE_LOCKED') and 'Закройте файл в Excel и повторите операцию' in message
assert not any(s in message for s in ('OTHER_ERROR','INVALID_XLSX','WinError 32'))
assert after['GUI_handler_returned'] and after['lock_still_live'] and after['no_partial_publication'] and after['old_good_unchanged']
assert retry['normal_generate_removed_own_old_pending'] and retry['no_harness_cleanup'] and retry['old_good_unchanged'] and retry['unrelated_unchanged']
for root in (E/'sandbox/output',E/'native_GUI/output'):
    assert not list(root.glob('.pending-*')) and not list(root.glob('.stage6-pending-*.json'))
for name in ('native_GUI_lock_0.png','native_GUI_lock_1.png','native_GUI_lock.txt'):assert (E/name).stat().st_size>1000
fixmanifest=load(FIX/'REVIEW_MANIFEST.json');assert all(sha(p)==h for p,h in fixmanifest['files'].items())
fixmodal=load(FIX/'native_GUI_message_while_locked.json');fixafter=load(FIX/'native_GUI_after_return_lock_still_held.json');fixretry=load(FIX/'native_GUI_release_cleanup_retry.json')
assert fixmodal['handle_live_during_modal'] and fixafter['handle_live_after_GUI_return'] and fixretry['pending_removed_by_product']
assert fixmodal['proof'][0]['winerror']==32 and fixmodal['message'].startswith('FILE_LOCKED')
save('tested_product_hashes_independent.json',{'BEFORE':b['product']['CURRENT'],'AFTER':product,'CURRENT':product,'BEFORE_EQUALS_AFTER_EQUALS_CURRENT':True,'pre_fix_baseline_separate':b['pre_fix_product']})
save('source_safety_final.json',{'rereview_guarded_count':len(current),'BEFORE_EQUALS_AFTER':True,'CURRENT':current,'expected_parallel_changes_from_fix_manifest':b['expected_parallel_changes'],'user_confirmation':b['parallel_changes_user_confirmed'],'parallel_changed_files_stable_during_rereview':True,'existing_pair':EXPECTED,'production_registry_attempts':[],'network_attempts':[],'network_tool_calls':0,'scope':'Only new rereview validation directory mutated. No product, real DEV, real source or prior output changes.'})
save('native_visual_observations.json',{'real_GUI_button_clicked_via_Sky':True,'actual_STOP_dialog_independently_read':True,'FILE_LOCKED':True,'actionable_Russian_instruction':True,'exact_pending_file_path_visible':True,'OTHER_ERROR_INVALID_XLSX_WinError32_not_in_user_message':True,'share0_lock_held_during_cleanup_modal_and_after_handler_return':True,'product_automatic_recovery_then_retry':True,'full_R1_visual_review_repeated':False})
save('findings.json',{'verdict':'PASS','S6-P1-R2':'CLOSED','R1':'PREVIOUSLY_CLOSED_SMOKE_PASS','counts':{'BLOCKER':0,'MAJOR':0,'MINOR':0},'findings':[],'required_fix':'NONE','ready_for_user_E2E':True})
parallel='\n'.join(r['path']+'\n  fix_manifest: '+r['fix_manifest_sha']+'\n  rereview_start/end: '+r['rereview_start_sha'] for r in b['expected_parallel_changes'])
producttext='\n'.join(p+'  '+h for p,h in product.items())
report=f'''1. VERDICT — PASS
S6-P1-R2 CLOSED. R1 ранее CLOSED/PASS; минимальный smoke PASS.
BLOCKER0 / MAJOR0 / MINOR0. Только R2 и связанная минимальная регрессия.

2. CURRENT PRODUCT VERIFIED
Все7 tested product files:final-tests BEFORE=AFTER=CURRENT;independent rereview BEFORE=AFTER=CURRENT.
{producttext}
Pre-fix core8dffda645c6370e2e284beaad44354b361ef9bd680ddfe190a74b92a404ada10.
Единственный новый product diff — core.py;gui.py и остальные6 product files неизменны.
Фактический diff равен product_changes.diff;BEFORE_core.py SHA связан с предыдущим independent manifest.
Из существующих определений изменены только FileOperationError и generate;добавлены8 pending helpers.
write_act,write_application,matching,preview,templates,date/climate/validation AST неизменны.

3. R2 PRIMARY ERROR PRESERVATION — PASS
Real primary FILE_LOCKED causeWinError32 и real cleanup FILE_LOCKED causeWinError32.
Тот же объект primary передан в error_message/GUI;код,путь,cause и инструкция сохранены.
Secondary только primary.cleanup_error/diagnostic JSON;GUI не показывает secondary WinError32.
Дополнительно6 категорий primary пережили injected secondary cleanup failure без подмены.
Реальное доказательство lock не заменено mock:real_lock_primary_proof.json + native_GUI_modal_lock_held.json.

4. REAL WINDOWS LOCK / GUI TEST — PASS
Здоровый pending APPLICATION перед lock совпал с SHA исходного шаблона.
CreateFile(GENERIC_READ,share=0) жив во время generate,cleanup,actual STOP/modal и после возврата GUI handler.
Кнопка «Сформировать TEST-пару» нажата через Sky. Фактическое сообщение:
{message}
OTHER_ERROR,INVALID_XLSX и WinError32 вместо основной причины отсутствуют.
Снимки:native_GUI_lock_0.png/native_GUI_lock_1.png;полный UI text:native_GUI_lock.txt.
Handle освобождён только после native_GUI_after_return_lock_held.json;preview invalidated,no partial publication.

5. DEFERRED PENDING CLEANUP — PASS
Durable record UUID,exact root/path,directory dev/inode и registered file dev/inode независимо сверены.
Повтор при живом lock => FILE_LOCKED;ID/size/mtime всех pending files и SHA ledger неизменны.
После release новый Python process без in-memory registry удалил собственный pending/ledger.
Unknown pending prefix,unknown XLSX inside owned pending,replaced expected-file identity,
replaced directory identity,tampered path и invalid UUID => PENDING_OWNERSHIP_STOP до удаления.
Во всех negative snapshots чужие/собственные файлы сохранены. Замены восстановлены только внутри review fixtures.
В отдельном настоящем GUI сценарии штатный generate сам очистил pending/ledger после release,
без ручной cleanup harness. Recursive product deletion отсутствует.

6. RETRY GENERATION — PASS
После FILE_LOCKED и release созданы ACT+APPLICATION+PAIR.json;receipt COMPLETE_TEST_PAIR,
единые9/2026,2026-10-05,act_id;receipt SHA соответствует обоим файлам.
Оба сценария:new-process recovery/retry и actual-GUI normal-generate automatic cleanup/retry PASS.
Старые успешные документы и отдельный user_document.xlsx не изменены;stale pending/ledger отсутствуют.
Pair evidence:native_GUI_release_product_cleanup_retry.json.

7. ERROR CLASSIFICATION — PASS
FILE_LOCKED:настоящий Windows32 и правильный GUI.
ACCESS_DENIED:Windows error5 unit,сообщение о правах доступа.
INVALID_XLSX:реально corrupt bytes/native Excel,distinct format error,cleanup завершён.
OUTPUT_EXISTS:OUTPUT_ALREADY_EXISTS_STOP,существующая пара не перезаписана.
INPUT_CHANGED:каждый из5 изменённых isolated inputs => INPUT_CHANGED_AFTER_PREVIEW_STOP.
OTHER_ERROR:валидный XLSX/generic non-format failure.
Для всех6 категорий primary object/code/message сохраняются при secondary cleanup failure.

8. R1 SMOKE — PASS
Native read-only saved isolated N3:4:6 visible,7:20 hidden,21 visible.
TOTAL3 шт./105.26kg/3.989m²;SUM C21/E21/G21 сохранены.
N17:полная пара,все17 исходных строк видимы;N18:CAPACITY STOP,без усечения/публикации.
Полный визуальный review/Print Preview не повторялся по scope пользователя.

9. SOURCE / TEST PAIR SAFETY
Обе ранее проверенные exact книги не пересоздавались и не открывались на запись:
ACT {EXPECTED[ACTNAME]}
APPLICATION {EXPECTED[APPNAME]}
Все{len(current)} текущих guarded fingerprints совпали BEFORE/AFTER rereview.
DEV ВЗМК_DEV.xlsx,УЗК_DEV.xlsm,Профили.xlsx,realcatalog,выбранная ВОМ9/2026,два АКЗ шаблона
и прочие защищённые файлы не изменены rereview. Product/реальные документы не правились.
ВАЖНО:в начале4 guarded файла отличались от fix manifest. Пользователь подтвердил:
«Да, изменения другой работы ожидаемы». Эти4 изменения явно выделены;не заявляется их неизменность
от старого fix manifest и они не приписываются R2. Их current hashes остались неизменными за rereview:
{parallel}
Stage1–5 сохранён на текущей подтверждённой baseline;полный Stage1–5 regression не выполнялся.
Production registry не использовался;network/web tools0;forbidden Python open/socket events0.

10. EVIDENCE INTEGRITY — PASS
81/81 fix evidence файлов SHA authenticated;previous-review manifest/BEFORE copy/current product binding проверены.
Reported Codex27/27 targeted PASS и16/16unit подтверждены по журналам и hashes.
Authors real GUI lock-through-return/recovery evidence прочитаны;новые независимые evidence выполнены отдельно.
Independent30/30 targeted PASS + current16/16unit PASS + actual native GUI lock/release/automatic cleanup/retry PASS.
Fix manifest-last проверен по составу/mtime. Этот REVIEW_MANIFEST.json записывается последним.
Старая baseline имеет4 объявленных parallel changes;current before/after hashes отдельно в source_safety_final.json.
Ни прежний PASS,ни authors journal не заменяют новые реальные lock/GUI проверки.

11. FINDINGS
BLOCKER0 / MAJOR0 / MINOR0. Незакрытых R2,непосредственных регрессий или критических evidence findings нет.

12. NOT_RUN
Real ACL mutation для ACCESS_DENIED (проверен error5 unit);physical print;full Stage1–5 regression;
повторный полный R1 visual/Print Preview;production/official issuance;Stage7;system packet capture;
index tools отсутствуют. Все обязательные real lock/actual GUI/fresh-process/ownership/retry сценарии выполнены.

13. REQUIRED FIX — NONE
Никакой product code не исправлялся. Mutation tests только в новой isolated validation directory.

14. READY FOR USER E2E — YES
S6-P1-R2 CLOSED. Готовность относится к user TEST E2E Stage6 Part1. STOP. Stage7 не начинался.
'''
(E/'REVIEW_REPORT.txt').write_text(report,encoding='utf-8')
files={str(p):sha(p) for p in sorted(E.rglob('*')) if p.is_file() and p.name!='REVIEW_MANIFEST.json'}
manifest={'review':'Stage6 FINAL TARGETED INDEPENDENT REREVIEW S6-P1-R2','created_at':datetime.now(timezone.utc).isoformat(),'scope':['S6-P1-R2','minimal related regression'],'verdict':'PASS','findings':{'BLOCKER':0,'MAJOR':0,'MINOR':0},'R2':'CLOSED','product':load(E/'tested_product_hashes_independent.json'),'expected_parallel_changes':b['expected_parallel_changes'],'parallel_changes_user_confirmed':b['parallel_changes_user_confirmed'],'existing_pair':EXPECTED,'reported_targeted':27,'independent_targeted':30,'independent_unit':16,'actual_GUI_persistent_lock_and_retry':'PASS','guarded_current':current,'files':files,'written_last':True,'required_fix':'NONE','ready_for_user_E2E':True,'stop_after_review':True}
save('REVIEW_MANIFEST.json',manifest)
assert all(sha(p)==h for p,h in files.items()) and all(sha(p)==h for p,h in current.items())
mp=E/'REVIEW_MANIFEST.json';assert all(Path(p).stat().st_mtime_ns<=mp.stat().st_mtime_ns for p in files)
print(json.dumps({'verdict':'PASS','R2':'CLOSED','findings':manifest['findings'],'independent_targeted':30,'unit':16,'actual_GUI':'PASS','guarded_before_after':len(current),'expected_parallel_changes':len(b['expected_parallel_changes']),'review_files':len(files),'manifest_last_verified':True,'manifest_sha256':sha(mp),'report':str(E/'REVIEW_REPORT.txt')},ensure_ascii=False,indent=2))
