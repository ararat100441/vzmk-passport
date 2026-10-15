from common import *
import tkinter as tk
import win32file,win32con,pywintypes
from contextlib import nullcontext
from decimal import Decimal
import copy,subprocess,traceback
product=load(E/'baseline.json')['product'];before={p:sha(p) for p in product};save('final_tests_product_BEFORE.json',before)
results=[]
def check(name,fn):
 try:
  value=fn();results.append({'name':name,'status':'PASS','detail':value});print('PASS',name,flush=True)
 except Exception as exc:results.append({'name':name,'status':'FAIL','error':str(exc),'traceback':traceback.format_exc()});raise
 finally:save('targeted_tests.json',{'results':results,'audit':AUDIT})
F,O=setup('sandbox_run2');sources=snapshot(F);old=snapshot(O/'OLD_GOOD');unrelated=sha(O/'unrelated.xlsx')
with isolated(F,O):
 p=c.preview('Гараж Казахстан','9/2026',F/'selected_VOM.xls',F/'TEST_catalog.xlsx')
 save('isolated_preview.json',{'seal':p.seal,'payload':p.payload})
 assert p.payload['totals']=={'quantity':Decimal(3),'mass_total':Decimal('105.26'),'area_total':Decimal('3.989')}
 handles=[];proof=[];primaries=[];original=c.native_open
 def locked(excel,path,*,read_only):
  if Path(path).name==APPNAME and not read_only:
   healthy=sha(path);assert healthy==sha(F/'APPLICATION_template.xlsx')
   handles.append(win32file.CreateFile(str(path),win32con.GENERIC_READ,0,None,win32con.OPEN_EXISTING,0,None))
   try:return original(excel,path,read_only=read_only)
   except Exception as exc:
    primaries.append(exc);proof.append({'path':str(path),'healthy_sha':healthy,'share':0,'code':getattr(exc,'code',None),'cause_winerror':getattr(exc.__cause__,'winerror',None)});raise
  return original(excel,path,read_only=read_only)
 root=tk.Tk();root.withdraw();gui=Launcher(root);gui.checked=p
 try:
  with patch.object(c,'native_open',side_effect=locked),patch('akz_stage6.gui.messagebox.showerror') as msg:gui.generate()
  message=msg.call_args.args[1];temp=Path(proof[0]['path']).parent;ledger=next(O.glob('.stage6-pending-*.json'));record=load(ledger)
  assert handles and all(win32file.GetFileType(h)==win32con.FILE_TYPE_DISK for h in handles)
  assert message.startswith('FILE_LOCKED') and 'Файл приложения АКЗ' in message and 'Закройте файл в Excel' in message and proof[0]['path'] in message and 'OTHER_ERROR' not in message
  assert primaries[0].code=='FILE_LOCKED' and primaries[0].cleanup_error.code=='FILE_LOCKED'
  assert proof[0]['cause_winerror']==32 and temp.is_dir() and set(f.name for f in temp.iterdir())=={ACTNAME,APPNAME}
  assert set(x.name for x in O.iterdir())=={'OLD_GOOD','unrelated.xlsx',temp.name,ledger.name}
  check('A_B_C_lock_held_through_GUI_return_primary_preserved_no_partial_pair',lambda:{'proof':proof,'message':message,'cleanup_error':str(primaries[0].cleanup_error),'handle_still_live':True,'ownership_record':record})
  private_locked={f.name:(f.lstat().st_ino,f.lstat().st_size,f.lstat().st_mtime_ns) for f in temp.iterdir()};ledger_sha=sha(ledger)
  check('retry_while_locked_returns_FILE_LOCKED_and_leaves_files_untouched',lambda:stop(lambda:c.generate(p,'STILL_LOCKED'),'FILE_LOCKED'))
  assert {f.name:(f.lstat().st_ino,f.lstat().st_size,f.lstat().st_mtime_ns) for f in temp.iterdir()}==private_locked and sha(ledger)==ledger_sha and not (O/'STILL_LOCKED').exists()
 finally:
  for h in handles:h.Close()
  root.destroy()
 private=snapshot(temp)
 # Unknown files must stop before any deletion. Fixtures removed only by this harness.
 foreign=temp/'unrelated_user.xlsx';foreign.write_bytes(b'FOREIGN PENDING CONTENT FIXTURE')
 check('F_unknown_file_inside_owned_pending_STOP_no_deletion',lambda:stop(c.recover_pending,'PENDING_OWNERSHIP_STOP'))
 assert foreign.read_bytes()==b'FOREIGN PENDING CONTENT FIXTURE' and snapshot(temp)|{}=={**private,str(foreign):sha(foreign)}
 foreign.unlink()
 # Substitute an expected filename with a new NTFS file identity, then restore original identity.
 app=temp/APPNAME;held=E/'held_original_application.xlsx';app.rename(held);shutil.copyfile(held,app)
 assert app.lstat().st_ino!=record['files'][APPNAME][1]
 replaced=snapshot(temp)
 check('F_replaced_expected_filename_identity_STOP_no_deletion',lambda:stop(c.recover_pending,'PENDING_OWNERSHIP_STOP'))
 assert snapshot(temp)==replaced
 app.unlink();held.rename(app)
 # Tampered root/path may not redirect cleanup to an unrelated directory.
 ledger_bytes=ledger.read_bytes();tampered=copy.deepcopy(record);tampered['path']=str(O/'OLD_GOOD');ledger.write_text(json.dumps(tampered),encoding='utf-8')
 check('F_tampered_ownership_path_STOP_no_deletion',lambda:stop(c.recover_pending,'PENDING_OWNERSHIP_STOP'))
 assert snapshot(O/'OLD_GOOD')==old and snapshot(temp)==private
 ledger.write_bytes(ledger_bytes)
 # Durable record is enough for a fresh Python process, without any in-memory registry.
 child=E/'recover_fresh_process.py';child.write_text('from common import *\nimport sys\nc.OUTPUT=Path(sys.argv[1])\nc.recover_pending()\nprint("RECOVERY_COMPLETE")\n',encoding='utf-8')
 completed=subprocess.run([sys.executable,'-B',str(child),str(O)],capture_output=True,text=True,encoding='utf-8',timeout=30)
 assert completed.returncode==0,completed.stderr
 assert not temp.exists() and not ledger.exists()
 check('D_pending_cleanup_after_release_fresh_process',lambda:{'owned_path':str(temp),'record':record,'child_stdout':completed.stdout,'removed':True})
 target=c.generate(p,'RETRY_RELEASED')
 assert load(target/'PAIR.json')['state']=='COMPLETE_TEST_PAIR' and set(x.name for x in target.iterdir())=={ACTNAME,APPNAME,'PAIR.json'}
 check('E_retry_complete_pair_no_stale_pending',lambda:{'target':str(target),'sha256':snapshot(target)})
 assert not list(O.glob('.pending-*')) and not list(O.glob('.stage6-pending-*.json'))
 # Smoke only: unchanged visibility/totals code, saved pair opened in native Excel.
 def smoke(target,p):
  with c.native_excel() as excel:
   c.validate_saved_pair(excel,target/ACTNAME,target/APPNAME,p)
   book=c.native_open(excel,target/APPNAME,read_only=True)
   try:
    s=book.Worksheets(1);n=len(p.payload['selected']);states={str(r):bool(s.Rows(r).Hidden) for r in range(1,22)}
    assert all(states[str(r)]==(4+n<=r<=20) for r in range(1,22))
    totals={col+'21':s.Range(col+'21').Value2 for col in ('C','E','G')}
    assert abs(Decimal(str(totals['C21']))-p.payload['totals']['quantity'])<Decimal('1e-8')
    assert abs(Decimal(str(totals['E21']))-p.payload['totals']['mass_total'])<Decimal('1e-8')
    assert abs(Decimal(str(totals['G21']))-p.payload['totals']['area_total'])<Decimal('1e-8')
    return {'N':n,'hidden':states,'totals':totals,'formulas':{col+'21':s.Range(col+'21').Formula for col in ('C','E','G')}}
   finally:book.Close(SaveChanges=False)
 check('J_R1_N3_smoke',lambda:smoke(target,p))
 p17=c.preview('Гараж Казахстан','9/2026',F/'SYNTHETIC_17.xlsx',F/'TEST_catalog.xlsx',synthetic=True)
 t17=c.generate(p17,'SMOKE_N17')
 check('J_R1_N17_smoke',lambda:smoke(t17,p17))
 check('J_R1_N18_capacity_STOP',lambda:stop(lambda:c.preview('Гараж Казахстан','9/2026',F/'SYNTHETIC_18.xlsx',F/'TEST_catalog.xlsx',synthetic=True),'TEMPLATE_CAPACITY_STOP'))
 unknown=O/'.pending-user-fixture';unknown.mkdir();(unknown/'USER.xlsx').write_bytes(b'NOT STAGE6')
 unknown_sha=snapshot(unknown)
 check('F_unknown_pending_prefix_is_not_ownership',lambda:stop(lambda:c.generate(p,'UNKNOWN_PENDING'),'PENDING_OWNERSHIP_STOP'))
 assert snapshot(unknown)==unknown_sha and not (O/'UNKNOWN_PENDING').exists()
 (unknown/'USER.xlsx').unlink();unknown.rmdir()
 # Actual malformed native input; cleanup succeeds and keeps format classification.
 def corrupt(excel,path,*,read_only):
  if Path(path).name==APPNAME and not read_only:Path(path).write_bytes(b'CORRUPT XLSX FIXTURE')
  return original(excel,path,read_only=read_only)
 with patch.object(c,'native_open',side_effect=corrupt):check('G_corrupt_XLSX_INVALID_XLSX_complete_cleanup',lambda:stop(lambda:c.generate(p,'CORRUPT'),'INVALID_XLSX'))
 assert not list(O.glob('.pending-*')) and not list(O.glob('.stage6-pending-*.json'))
 check('H_OUTPUT_EXISTS_no_overwrite',lambda:stop(lambda:c.generate(p,'RETRY_RELEASED'),'OUTPUT_ALREADY_EXISTS_STOP'))
 for path in (c.REGISTRY,c.ACT,c.APPLICATION,F/'TEST_catalog.xlsx',F/'selected_VOM.xls'):
  data=path.read_bytes()
  try:
   path.write_bytes(data+b'ISOLATED CHANGED INPUT')
   check('I_INPUT_CHANGED_'+path.name,lambda:stop(lambda:c.generate(p,'CHANGED'),'INPUT_CHANGED_AFTER_PREVIEW_STOP'))
  finally:path.write_bytes(data)
 denied=pywintypes.error(5,'CreateFile','Access denied')
 with patch('win32file.CreateFile',side_effect=denied):check('ACCESS_DENIED_distinct_unit',lambda:stop(lambda:c.probe_file_access(c.APPLICATION,read_only=False),'ACCESS_DENIED'))
 check('OTHER_ERROR_only_generic_failure',lambda:stop(lambda:c.excel_file_failure(c.APPLICATION,RuntimeError('generic'),read_only=True,operation='test'),'OTHER_ERROR'))
 # Every primary category survives an injected secondary cleanup exception unchanged.
 for code in ('FILE_LOCKED','ACCESS_DENIED','INVALID_XLSX','OTHER_ERROR','OUTPUT_ALREADY_EXISTS_STOP','INPUT_CHANGED_AFTER_PREVIEW_STOP'):
  primary=c.FileOperationError(code,c.APPLICATION,'test') if code in ('FILE_LOCKED','ACCESS_DENIED','INVALID_XLSX','OTHER_ERROR') else c.Stage6Error(code)
  cleanup=PermissionError(13,'injected secondary cleanup failure');cleanup.winerror=32
  with patch.object(c,'native_excel',side_effect=lambda:nullcontext(None)),patch.object(c,'write_act',side_effect=primary),patch.object(c,'cleanup_pending',side_effect=cleanup):
   try:c.generate(p,'PRIMARY_'+code)
   except Exception as exc:assert exc is primary and exc.cleanup_error is cleanup
   else:raise AssertionError('Expected primary exception')
  check('B_primary_preserved_through_cleanup_failure_'+code,lambda:{'primary_type':type(primary).__name__,'message':str(primary),'cleanup_type':type(cleanup).__name__})
  c.recover_pending()
 assert snapshot(F)==sources and snapshot(O/'OLD_GOOD')==old and sha(O/'unrelated.xlsx')==unrelated
 assert not list(O.glob('.pending-*')) and not list(O.glob('.stage6-pending-*.json'))
 check('K_source_old_pair_unrelated_files_unchanged',lambda:{'sources':sources,'old_good':old,'unrelated_sha256':unrelated})
after={p:sha(p) for p in product};assert before==after
save('final_tests_product_AFTER.json',after);save('audit.json',AUDIT)
print('ALL TARGETED PASS',len(results),flush=True)
