from review_common import *
from datetime import datetime,timezone
b=load(E/'baseline_current.json');fix=load(FIX/'EVIDENCE_MANIFEST.json');old=load(FIX/'baseline.json')
current={p:sha(p) for p in b['guarded']}
assert current==b['guarded'],'REREVIEW SOURCE SAFETY STOP'
delta=[p for p,h in old['guarded'].items() if current[p]!=h]
expected_changed=[str(ROOT/'akz_stage6'/n) for n in ('core.py','gui.py')]
assert set(delta)==set(expected_changed)
previous=load(old['review_manifest'])
previous_all={**previous['guarded_current'],**previous['files'],old['review_manifest']:sha(old['review_manifest'])}
assert old['guarded']==previous_all
assert all(old['product'][p]==previous['guarded_current'][p] for p in old['product'])
assert load(PAIR/'PAIR.json')['sha256']=={n:h for n,h in EXPECTED.items() if n.endswith('.xlsx')}
assert load(PAIR/'PAIR.json')['state']=='COMPLETE_TEST_PAIR'
actualdiff=''
for name in ('core.py','gui.py'):
    before=(FIX/('BEFORE_'+name)).read_text('utf-8');after=(ROOT/'akz_stage6'/name).read_text('utf-8')
    actualdiff+=''.join(difflib.unified_diff(before.splitlines(True),after.splitlines(True),fromfile='reviewed/'+name,tofile='current/'+name))
assert actualdiff==(FIX/'product_changes.diff').read_text('utf-8')
oldtree=ast.parse((FIX/'BEFORE_core.py').read_text('utf-8'));newtree=ast.parse((ROOT/'akz_stage6/core.py').read_text('utf-8'))
dump=lambda n:ast.dump(n,include_attributes=False)
oldfn={n.name:n for n in oldtree.body if isinstance(n,ast.FunctionDef)}
newfn={n.name:copy.deepcopy(n) for n in newtree.body if isinstance(n,ast.FunctionDef)}
class RestoreSave(ast.NodeTransformer):
    def visit_Call(self,n):
        n=self.generic_visit(n)
        if isinstance(n.func,ast.Name) and n.func.id=='native_save':
            assert [dump(x) for x in n.args]==[dump(ast.Name(id='b',ctx=ast.Load())),dump(ast.Name(id='path',ctx=ast.Load()))]
            n.func=ast.Attribute(value=ast.Name(id='b',ctx=ast.Load()),attr='Save',ctx=ast.Load());n.args=[]
        return n
assert dump(RestoreSave().visit(newfn['write_act']))==dump(oldfn['write_act'])
remove=ast.parse("count = len(data['selected'])\nrequire(1 <= count <= 17, 'TEMPLATE_CAPACITY_STOP: Part 1 поддерживает до 17 строк №0')\ns.Rows('1:21').Hidden = False\nif count < 17:\n    s.Rows(f'{4+count}:20').Hidden = True").body
class RemoveVisibility(ast.NodeTransformer):
    def __init__(self):self.removed=0
    def visit(self,n):
        if dump(n) in {dump(x) for x in remove}:
            self.removed+=1;return None
        return super().visit(n)
visibility=RemoveVisibility();app=visibility.visit(newfn['write_application'])
assert visibility.removed==4
assert dump(RestoreSave().visit(app))==dump(oldfn['write_application'])
v=newfn['validate_saved_pair']
class RemoveHiddenRequire(ast.NodeTransformer):
    def visit_Expr(self,n):
        if isinstance(n.value,ast.Call) and any(isinstance(a,ast.Constant) and a.value=='SAVED_APPLICATION_ROW_VISIBILITY_STOP' for a in n.value.args):return None
        return self.generic_visit(n)
assert dump(RemoveHiddenRequire().visit(v))==dump(oldfn['validate_saved_pair'])
n=newfn['native_open'];o=oldfn['native_open']
assert dump(n.body[1].body[0])==dump(o.body[0]) and [dump(x) for x in n.body[2:]]==[dump(x) for x in o.body[1:]]
shared=lambda t:[dump(n) for n in t.body if not isinstance(n,(ast.FunctionDef,ast.ClassDef)) and not (isinstance(n,ast.ImportFrom) and n.module in ('zipfile','xml.etree'))]
assert shared(oldtree)==shared(newtree)
proof={'actual_diff_matches_reported_byte_for_byte_after_text_newline_normalization':True,'old_before_copies_bound_to_prior_review':True,'mathematics_template_identity_number_date_climate_unchanged_AST':True,'generate_atomic_publication_and_cleanup_unchanged_AST':True,'GUI_only_error_message_routing_changed':True}
save('semantic_scope_independent.json',proof)
tests=load(E/'independent_targeted.json');assert tests['summary']=={'PASS':15,'FAIL':2} and tests['units']=={'run':16,'pass':True}
handler=load(E/'persistent_lock_GUI_handler.json');native=load(E/'native_GUI_lock_after_OK.json')
assert handler['proof'][0]['cause_winerror']==32 and native['proof'][0]['cause_winerror']==32
assert native['proof'][0]['classified_code']=='FILE_LOCKED' and native['messages'][0]['message'].startswith('OTHER_ERROR: [WinError 32]')
assert native['pending'] and not native['published'] and native['old_good_unchanged']
assert load(E/'native_GUI_release_retry.json')['old_good_unchanged']
assert not list((E/'sandbox/output').glob('.pending-*')) and not list((E/'native_GUI_output').glob('.pending-*'))
assert load(E/'native_exact_close.json')['hashes_after']==EXPECTED
for name in ('native_application_print_preview.png','native_act_print_preview.png','native_GUI_lock_1.png'):
    assert (E/name).stat().st_size>10000
observations={'independent_visual_review':True,'exact_read_only_pair':True,'normal_Open_CorruptLoad_0':True,'repair_recovery_or_restored_title_observed':False,'application_print_preview':{'pages':1,'three_marks_followed_immediately_by_total':True,'table_not_clipped':True,'borders_preserved':True},'act_print_preview':{'pages':1,'number':'9/2026','date':'05.10.2026','TEST_marking_visible':True},'closed_without_save_hash_unchanged':True,'GUI_lock_visual':{'button_clicked_in_real_GUI':True,'exclusive_CreateFile_share':0,'lock_held_through_GUI_error_and_OK':True,'displayed_code':'OTHER_ERROR','expected_code':'FILE_LOCKED / ACCESS_DENIED','close_file_and_retry_instruction_present':False},'UI_tool_limitations':'Initial Excel window snapshots/activation conflicted with minimized-window/user-input fence. Closed own read-only instance without saving, reopened exact pair; completed both Print Previews with supported Sky API. Modal OK element lookup failed; fresh observation and Return closed own STOP dialog. No product mutation.'}
save('native_visual_observations.json',observations)
finding={'id':'S6-P1-R2','severity':'MINOR','status':'OPEN','file':str(ROOT/'akz_stage6/core.py'),'line':524,'title':'Удерживаемая блокировка маскируется ошибкой rollback и оставляет pending-файл','trigger':'Healthy APPLICATION pending copy is held by actual Windows CreateFile GENERIC_READ share=0 until after GUI returns / modal OK.','expected':'FILE_LOCKED/ACCESS_DENIED, exact path, actionable Russian instruction, controlled rollback and successful retry after release.','actual':'native_open correctly raises FileOperationError(FILE_LOCKED), cause winerror32; generate finally shutil.rmtree fails on same live lock, replaces original exception; real GUI displays OTHER_ERROR WinError32 without close-and-retry instruction; private pending directory retains APPLICATION. No target pair published; prior valid pair unchanged. Release then retry succeeds.','evidence':['persistent_lock_GUI_handler.json','persistent_lock_after_release.json','native_GUI_lock_after_OK.json','native_GUI_lock.txt','native_GUI_lock_1.png','native_GUI_release_retry.json'],'reported_test_gap':'Both fix drivers close the OS handle in native_open wrapper finally before generate executes cleanup; their 16/16 result does not cover persistent external lock.'}
save('findings.json',{'verdict':'FAIL','R1':'CLOSED','R2':'OPEN','counts':{'BLOCKER':0,'MAJOR':0,'MINOR':1},'findings':[finding]})
safety={'guarded_rereview_count':len(current),'all_current_equals_rereview_baseline':True,'prior_review_guard_count':len(old['guarded']),'prior_review_unchanged_excluding_two_authorized_product_fixes':len(old['guarded'])-2,'intentional_before_current_changes':delta,'product_CURRENT':fix['product']['CURRENT'],'old_pair_hashes':{p:h for p,h in current.items() if 'stage6_akz_part1_20261009_085822_release3_TEST_9_2026' in p},'new_pair':EXPECTED,'fix_evidence_files_verified':len(fix['files']),'production_open_attempts':tests['audit']['production_attempts'],'network_attempts':tests['audit']['network_attempts'],'network_tool_calls':0,'scope':'Mutations restricted to this newly created rereview validation directory. No real DEV/source/old output/new exact output/product edits. Python audit is not system packet capture.'}
save('source_safety_final.json',safety)
hashtext='\n'.join(p+'  '+h for p,h in fix['product']['CURRENT'].items())
newtext='\n'.join(n+'  '+h for n,h in EXPECTED.items())
report=f'''1. VERDICT — FAIL
R1 CLOSED. R2 OPEN. Один незакрытый S6-P1-R2; новых замечаний вне scope нет.

2. CURRENT ARTIFACT VERIFIED
Фактический diff равен product_changes.diff. BEFORE связан с предыдущим review и BEFORE copies.
AFTER=CURRENT для всех 7 product/launcher/docs/test файлов; повторно CURRENT неизменен.
{hashtext}
Exact новая пара (не заменялась и не пересоздавалась):
{newtext}
Полный BEFORE/AFTER/CURRENT: baseline_current.json; итог CURRENT: source_safety_final.json.

3. R1 REREVIEW — PASS
В exact APP: 4:6 VISIBLE, 7:20 HIDDEN, 21 VISIBLE.
Формулы C21/E21/G21 = SUM(C4:C20)/SUM(E4:E20)/SUM(G4:G20).
Итоги 3 / 105.26 кг / 3.989 м². Cached E21 105.26000000000002 — только IEEE transport noise, формат 105,26; без изменения округления.
Все значения/формулы/границы/шрифты/объединения/ширины/заголовки/print area совпали с прежней парой.
Native Print Preview приложения: 3 строки сразу перед Итого; 1 страница, обрезания нет.
Isolated N1: row4 visible,5:20 hidden,21 visible; N3:4:6 visible,7:20 hidden,21 visible;
N17:4:20 visible,21 visible,17 исходных строк сохранены; N18:TEMPLATE_CAPACITY_STOP,без усечения/публикации.
Проверены завершённые пары/receipt и сохранённые XLSX. Числовой код AST неизменен.

4. R2 REREVIEW — FAIL
Реальная блокировка APPLICATION CreateFile(GENERIC_READ,share=0), здоровый SHA до lock подтверждён.
core.native_open корректно выдаёт FILE_LOCKED с cause winerror32.
Но lock удерживается до сообщения GUI: generate finally/core.py:524 не удаляет locked pending XLSX,
shutil.rmtree выбрасывает PermissionError WinError32 и подменяет FileOperationError.
Кнопка реального GUI нажата через Sky; фактический STOP: OTHER_ERROR: [WinError 32] ... занят другим процессом.
Путь присутствует, сообщения о corrupt XLSX нет, но actionable инструкция закрыть файл и повторить потеряна.
Остаётся private .pending-.../APPLICATION; полный rollback FAIL. Неполная пара не опубликована.
Старая valid pair неизменна. После release повторная generation COMPLETE_TEST_PAIR — PASS.
Наблюдение воспроизведено дважды: GUI handler и настоящий GUI/modal screenshot.
Остатки удалены только review harness после release и фиксации доказательств; продукт их сам не удалил.
Differentiation: real corrupt bytes/native Excel => INVALID_XLSX + полный rollback PASS;
existing => OUTPUT_ALREADY_EXISTS_STOP + no overwrite PASS;
каждый из 5 isolated inputs changed after preview => INPUT_CHANGED_AFTER_PREVIEW_STOP PASS;
Windows error5 unit => ACCESS_DENIED PASS; valid XLSX/generic error => OTHER_ERROR PASS.

5. THREE-MARK REGRESSION — PASS
Native исходная выбранная ВОМ через изолированную точную копию:104 visible/397 hidden; порядок сохранён.
Бп2-5,row40:1 шт.; unit/total mass37.20kg;unit/total area1.412m².
Бп2-18,row53:1 шт.; unit/total mass31.32kg;unit/total area1.184m².
Бп2-19,row54:1 шт.; unit/total mass36.74kg;unit/total area1.393m².
TOTAL3 /105.26kg /3.989m². Exact APP содержит эти же значения.
Pirocor Protect057;RAL7004;1 слой;80µm;9/2026;05.10.2026.
ТЕСТ/ЧЕРНОВИК сохранён;24 factual cells только MISSING/TEST placeholders,исторические измерения не выданы за новые.

6. NATIVE EXCEL VALIDATION — PASS
Excel16.0; exact новые ACT+APP ReadOnly=True,UpdateLinks0,CorruptLoad0,DisplayAlerts=True.
No repair/recovery/[восстановлен] observed. По одному ожидаемому листу и одной печатной странице.
Оба Print Preview независимо осмотрены и сохранены PNG/TXT. Не нажималась физическая Печать.
Закрыто SaveChanges=False; SHA двух XLSX и PAIR.json прежние.

7. SOURCE / OLD ARTIFACT SAFETY — PASS
Все {len(current)} текущих guarded fingerprints совпали до/после rereview.
От предыдущего review:381/381 unchanged,только два разрешённых изменения core/gui.
Включены Stage1–5,DEV ВЗМК_DEV.xlsx,УЗК_DEV.xlsm,Профили.xlsx,realcatalog,ВОМ,шаблоны,old evidence/outputs.
Old ACT1a52736427e1529420723ff85f3bfd92340865ece895173ffaaa0dadb48ba06f;
old APP4cd657f84311512255e6a6e39b492c227f18df3ab8b5ab03eb34400f165d9c6d — unchanged.
Production registry не открывался. Web/network tools0; Python open/socket audit0 forbidden attempts.
Никаких product/real DEV/source/output edits; mutations только внутри этого rereview validation folder.

8. EVIDENCE INTEGRITY
72/72 fix files verified по manifest;383/383 guarded_current совпали;BEFORE copies/prior review manifest hashes совпали.
Fix manifest-last подтверждён по составу/mtime. Новый REVIEW_MANIFEST.json записывается последним.
Reported targeted16/16 и unit16/16 журналы подлинны по SHA, но lock fixtures освобождают lock перед cleanup.
Independent17 targeted assertions:15 PASS,2 FAIL;оба FAIL относятся к одному S6-P1-R2.
Independent current unit suite16/16 PASS. Native exact pair дополнительно проверена отдельным read-only driver/UI.
Review harness initially compared cached floating point totals exactly; adjusted only harness to 1e-8 transport tolerance.
UI helper initial activation/cache issues resolved by closing own read-only Excel and reopening exact pair; previews completed.
Ни старый PASS,ни original test16/16 не заменяют нынешний независимый lock result.

9. FINDINGS
BLOCKER0 / MAJOR0 / MINOR1.
S6-P1-R2 OPEN (MINOR):cleanup при удерживаемом lock маскирует FILE_LOCKED как OTHER_ERROR и оставляет pending.
Source:akz_stage6/core.py:522-524;GUI route:akz_stage6/gui.py generate except.
Evidence:persistent_lock_GUI_handler.json;native_GUI_lock_after_OK.json;native_GUI_lock_1.png.

10. NOT_RUN
Real ACL mutation для ACCESS_DENIED (проверен error5 unit);physical print;full Stage1–5 regression;
production access/output;official issuance;Stage7;system packet capture;repository index tools unavailable.
Все обязательные native Open/Print Preview и real lock/GUI/retry сценарии выполнены.

11. REQUIRED FIX
Закрыть только S6-P1-R2:сохранить исходную typed/actionable lock-ошибку при сбое cleanup;
обеспечить управляемый rollback и очистку остатка после release,без публикации неполной пары;
повторить regression с OS handle,удерживаемым до возврата GUI,а не закрываемым в native_open finally.
В этом rereview код не исправлялся.

12. READY FOR USER E2E — NO
R1 принят;R2 требует исправления и targeted rereview. STOP. Stage7 не начинался.
'''
(E/'REVIEW_REPORT.txt').write_text(report,encoding='utf-8')
files={str(p):sha(p) for p in sorted(E.rglob('*')) if p.is_file() and p.name!='REVIEW_MANIFEST.json'}
manifest={'review':'Stage6 targeted independent R1/R2 rereview','created_at':datetime.now(timezone.utc).isoformat(),'scope':['S6-P1-R1','S6-P1-R2'],'verdict':'FAIL','findings':{'BLOCKER':0,'MAJOR':0,'MINOR':1},'product':fix['product'],'new_pair':EXPECTED,'targeted_independent':tests['summary'],'unit_tests':tests['units'],'guarded_current':current,'files':files,'written_last':True,'stop_after_review':True}
save('REVIEW_MANIFEST.json',manifest)
assert all(sha(p)==h for p,h in files.items()) and all(sha(p)==h for p,h in current.items())
manifest_path=E/'REVIEW_MANIFEST.json'
assert all(Path(p).stat().st_mtime_ns<=manifest_path.stat().st_mtime_ns for p in files)
print(json.dumps({'verdict':'FAIL','R1':'CLOSED','R2':'OPEN','findings':manifest['findings'],'guarded_count':len(current),'review_files':len(files),'manifest_sha256':sha(manifest_path),'manifest_last_verified':True,'report':str(E/'REVIEW_REPORT.txt')},ensure_ascii=False,indent=2))
