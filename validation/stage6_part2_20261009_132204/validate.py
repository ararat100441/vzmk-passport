from pathlib import Path
import sys,json,copy,subprocess,traceback,os
from decimal import Decimal
from unittest.mock import patch
ROOT=Path(__file__).resolve().parents[2];OUT=Path(__file__).resolve().parent
sys.path.insert(0,str(ROOT))
from akz_stage6 import core as c,part2 as p2
from openpyxl import load_workbook
from akz_stage6.gui import Launcher
import tkinter as tk
import win32file,win32con
def dump(n,x):(OUT/n).write_text(json.dumps(x,ensure_ascii=False,indent=2,default=str),encoding='utf-8')
results=[];audit={'network_events':[],'forbidden_production_events':[]}
forbidden={str(ROOT.parents[1]/'ВЗМК.xlsx').casefold(),str(ROOT.parents[1]/'УЗК.xlsm').casefold()}
def audit_hook(event,args):
 if event in ('socket.connect','socket.getaddrinfo'):
  audit['network_events'].append(event);raise RuntimeError('NETWORK_FORBIDDEN')
 if event=='open' and isinstance(args[0],(str,bytes)) and str(args[0]).casefold() in forbidden:
  audit['forbidden_production_events'].append(str(args[0]));raise RuntimeError('PRODUCTION_FORBIDDEN')
sys.addaudithook(audit_hook)
def check(name,fn):
 try:
  detail=fn();results.append({'name':name,'status':'PASS','detail':detail});print('PASS',name,flush=True)
 except BaseException as exc:
  results.append({'name':name,'status':'FAIL','error':str(exc),'traceback':traceback.format_exc()});raise
 finally:dump('targeted_validation.json',{'results':results,'audit':audit})
def yes(value):assert value;return True
def stop(fn,code):
 try:fn()
 except Exception as exc:
  assert code in str(exc),str(exc);return {'code':code,'message':str(exc)}
 raise AssertionError('Expected STOP '+code)
product=[ROOT/'akz_stage6/core.py',ROOT/'akz_stage6/gui.py',ROOT/'akz_stage6/part2.py',ROOT/'tests/test_stage6.py',ROOT/'tests/test_stage6_part2.py',ROOT/'STAGE6_DEV.md']
dump('tested_product_before.json',{str(p):c.sha256(p) for p in product})
catalog=c.FIXTURES/'stage6_part2_20261009_132204/TEST_Справочник марок.xlsx'
vom=c.PROJECT/'Паспорт/9 от 06.10.2026/ВОМ отгрузочный счет 183  (5).xls'
batch=p2.preview_groups('Гараж Казахстан','9/2026',vom,catalog)
generation=json.loads((OUT/'real_generation.json').read_text(encoding='utf-8'))
def real_pairs():
 evidence=[]
 for n in (1,2,3):
  p=p2.group_preview(batch,n);target=Path(generation['groups'][str(n)]['path']);names=p2.filenames('9/2026',n)
  hashes={name:c.sha256(target/name) for name in names[:2]}
  with c.native_excel() as app:native=p2.validate_saved_pair(app,target/names[0],target/names[1],p)
  assert hashes=={name:c.sha256(target/name) for name in names[:2]}
  receipt=json.loads((target/'PAIR.json').read_text(encoding='utf-8'));assert receipt['sha256']==hashes
  evidence.append({'variant':n,'path':str(target),'sha256':hashes,'totals':p.payload['totals'],'native':native})
 return evidence
check('all_six_real_outputs_current_product_readonly_full_calculation',real_pairs)
check('one_daily_missing_climate_all_groups',lambda:yes(len(batch.payload['climate_days'])==1 and len({c.digest(batch.payload['groups'][str(n)]['payload']['climate']) for n in (1,2,3)})==1))
def units():
 expected=[(1,'Бп2-5',40,Decimal('37.2'),Decimal('1.412')),(2,'Бп2-18',53,Decimal('31.32'),Decimal('1.184')),(3,'Бп2-19',54,Decimal('36.74'),Decimal('1.393'))]
 for n,mark,rn,mass,area in expected:
  selected=p2.group_preview(batch,n).payload['selected'];assert len(selected)==1
  r=selected[0];assert (r['mark'],r['row'],r['quantity'],r['mass_unit'],r['area_unit'])==(mark,rn,Decimal(1),mass,area)
 return expected
check('real_VOM_row_provenance_unit_values_no_cross_group',units)
sandbox=OUT/'sandbox_final2';sandbox.mkdir(exist_ok=False)
with patch.object(p2,'OUTPUT',sandbox/'part2'),patch.object(c,'OUTPUT',sandbox/'part0'):
 for n,cap in ((1,61),(2,47),(3,94)):
  def boundary(n=n,cap=cap):
   base=p2.group_preview(batch,n).payload;r=base['selected'][0]
   data=copy.deepcopy(base);data['synthetic']=True;data['act_number']='SYNTHETIC-TEST 9/2026-'+str(n)
   data['selected']=[{**r,'row':10000+i} for i in range(cap)]
   data['vom']={'path':'SYNTHETIC_BOUNDARY_IN_MEMORY','rows':data['selected']}
   data['totals']={f:sum((r[f] for r in data['selected']),Decimal(0)) for f in ('quantity','mass_total','area_total')}
   p=c.Preview(data,c.digest(data));target=p2.generate(p,f'SYNTHETIC_CAP_{n}_{cap}')
   bad=copy.deepcopy(data);bad['selected'].append(dict(r));bad=c.Preview(bad,c.digest(bad))
   detail=stop(lambda:p2.generate(bad,f'CAP_OVER_{n}'),'CAPACITY_STOP')
   assert not (p2.OUTPUT/f'CAP_OVER_{n}').exists()
   return {'capacity':cap,'overflow':cap+1,'full_native_pair':str(target),'overflow_stop':detail}
  check(f'native_capacity_{n}_{cap}_{cap+1}',boundary)
 p=p2.group_preview(batch,2)
 def overwrite():
  name='EXISTING';target=p2.OUTPUT/name;target.mkdir();(target/'USER.txt').write_text('KEEP',encoding='utf-8')
  detail=stop(lambda:p2.generate(p,name),'OUTPUT_ALREADY_EXISTS_STOP');assert (target/'USER.txt').read_text()=='KEEP';return detail
 check('no_overwrite_existing_folder_or_user_file',overwrite)
 def rollback():
  with patch.object(p2,'write_application',side_effect=c.Stage6Error('INJECTED_APPLICATION_FAILURE')):
   detail=stop(lambda:p2.generate(p,'ROLLBACK'),'INJECTED_APPLICATION_FAILURE')
  assert not (p2.OUTPUT/'ROLLBACK').exists() and not list(p2.OUTPUT.glob('.pending-*')) and not list(p2.OUTPUT.glob('.part2-pending-*'))
  return detail
 check('pair_rollback_after_actual_act_save',rollback)
 def changed_inputs():
  details=[]
  for i,source in enumerate(p.payload['sources']):
   isolated=sandbox/f'input_{i}.bin';isolated.write_bytes(Path(source).read_bytes())
   data=copy.deepcopy(p.payload);data['sources'][str(isolated)]=c.sha256(isolated);preview=c.Preview(data,c.digest(data))
   isolated.write_bytes(b'CHANGED_ISOLATED_INPUT')
   details.append(stop(lambda:p2.generate(preview,f'CHANGED_{i}'),'INPUT_CHANGED_AFTER_PREVIEW_STOP'))
  return details
 check('all_five_changed_inputs_stop_isolated_copies',changed_inputs)
 def lock_recovery(module,p,name):
  original=c.native_open;handles=[];primary=[];proof={}
  def locked(app,path,*,read_only):
   if 'Приложение' in Path(path).name and not read_only:
    healthy=c.sha256(path)
    handles.append(win32file.CreateFile(str(path),win32con.GENERIC_READ,0,None,win32con.OPEN_EXISTING,0,None))
    expected=p.payload['sources'][str(module.layout(2).application)] if module is p2 else c.PINS[c.APPLICATION]
    assert healthy==expected
    proof['healthy_sha']=healthy
    try:return original(app,path,read_only=read_only)
    except BaseException as exc:primary.append(exc);raise
   return original(app,path,read_only=read_only)
  try:
   with patch.object(c,'native_open',side_effect=locked):
    try:module.generate(p,name)
    except c.FileOperationError as exc:
     assert exc is primary[0] and exc.code=='FILE_LOCKED' and exc.__cause__.winerror==32
     assert exc.cleanup_error.code=='FILE_LOCKED' and exc.cleanup_error.__cause__.winerror==32
     proof.update(primary=str(exc),cleanup=str(exc.cleanup_error),same_primary_object=True,live_handle_after_return=win32file.GetFileType(handles[0])==win32con.FILE_TYPE_DISK)
    else:raise AssertionError('LOCK DID NOT STOP')
   root=module.OUTPUT;pending=next(root.glob('.pending-*'));ledger=next(root.glob('.part2-pending-*' if module is p2 else '.stage6-pending-*'))
   stat_before={x.name:[x.stat().st_ino,x.stat().st_size,x.stat().st_mtime_ns] for x in pending.iterdir()};ledger_before=c.sha256(ledger)
   stop(module.recover_pending,'FILE_LOCKED')
   assert stat_before=={x.name:[x.stat().st_ino,x.stat().st_size,x.stat().st_mtime_ns] for x in pending.iterdir()} and ledger_before==c.sha256(ledger)
   unknown=root/'.pending-foreign';unknown.mkdir();(unknown/'USER.txt').write_bytes(b'KEEP')
   stop(module.recover_pending,'PENDING_OWNERSHIP_STOP');assert (unknown/'USER.txt').read_bytes()==b'KEEP'
   (unknown/'USER.txt').unlink();unknown.rmdir()
   inside=pending/'USER.txt';inside.write_bytes(b'KEEP');stop(module.recover_pending,'PENDING_OWNERSHIP_STOP');assert inside.read_bytes()==b'KEEP';inside.unlink()
   act=next(x for x in pending.iterdir() if x.suffix=='.xlsx' and 'Приложение' not in x.name)
   backup=act.with_suffix('.held');assert backup.resolve().is_relative_to(sandbox.resolve())
   act.rename(backup);act.write_bytes(backup.read_bytes())
   stop(module.recover_pending,'PENDING_OWNERSHIP_STOP');act.unlink();backup.rename(act)
  finally:
   for h in handles:h.Close()
  child=OUT/('recover_part2.py' if module is p2 else 'recover_part0.py')
  child.write_text('from pathlib import Path\nimport sys\nfrom akz_stage6 import '+('part2' if module is p2 else 'core')+' as m\nm.OUTPUT=Path(sys.argv[1])\nm.recover_pending()\n',encoding='utf-8')
  completed=subprocess.run([sys.executable,'-B',str(child),str(root)],cwd=ROOT,env={**os.environ,'PYTHONPATH':str(ROOT)},capture_output=True,text=True,encoding='utf-8')
  assert completed.returncode==0,completed.stderr
  assert not pending.exists() and not ledger.exists() and not (root/name).exists()
  target=module.generate(p,name);proof['retry_target']=str(target);proof['fresh_process_recovery']=True
  return proof
 check('part2_real_Windows_lock_primary_cleanup_unknown_replacement_fresh_process_retry',lambda:lock_recovery(p2,p,'LOCKED_RETRY'))
 old_fixture=c.FIXTURES/'stage6_akz_part1_20261009_085822_release3/TEST_Справочник марок.xlsx'
 zero=c.preview('Гараж Казахстан','9/2026',vom,old_fixture)
 def r1():
  target=c.generate(zero,'R1_N3_NEW_ISOLATED')
  with c.native_excel() as app:
   b=c.native_open(app,target/'ТЕСТ Приложение к АКЗ №0.xlsx',read_only=True)
   try:
    s=b.Worksheets('Приложение АКЗ');assert [r for r in range(1,22) if s.Rows(r).Hidden]==list(range(7,21));assert not s.Rows(21).Hidden
    for col,expected in zip('CEG',('3','105.26','3.989')):
     assert abs(Decimal(str(s.Range(col+'21').Value2))-Decimal(expected))<=Decimal('0.00000001')
   finally:b.Close(SaveChanges=False)
  data=copy.deepcopy(zero.payload);data['selected']=(data['selected']*6)[:17];data['synthetic']=True;data['act_number']='SYNTHETIC-TEST 9/2026'
  data['totals']={f:sum((r[f] for r in data['selected']),Decimal(0)) for f in ('quantity','mass_total','area_total')}
  full=c.generate(c.Preview(data,c.digest(data)),'R1_N17_SYNTHETIC')
  rows=[zero.payload['selected'][0]]*18
  stop(lambda:c.matching({'rows':rows},{rows[0]['mark']:{'assignment':0,'row':2}}),'TEMPLATE_CAPACITY_STOP')
  return {'N3':str(target),'N17':str(full),'N18':'CAPACITY_STOP','accepted_old_outputs':'NOT_OVERWRITTEN'}
 check('R1_native_N3_N17_N18_regression',r1)
 check('R2_zero_real_lock_cleanup_recovery_retry',lambda:lock_recovery(c,zero,'R2_ZERO_RETRY'))
 def gui_zero_lock():
  root=tk.Tk();root.withdraw();gui=Launcher(root)
  zero_batch=p2.preview_groups('Гараж Казахстан','9/2026',vom,old_fixture)
  gui.checked=zero_batch;gui.variant_vars[0].set(True);gui.passport.set('9/2026');gui.checked=zero_batch
  original=c.native_open;handles=[];primary=[];seen=[];message=[]
  def locked(app,path,*,read_only):
   if Path(path).name=='ТЕСТ Приложение к АКЗ №0.xlsx' and not read_only:
    handles.append(win32file.CreateFile(str(path),win32con.GENERIC_READ,0,None,win32con.OPEN_EXISTING,0,None))
    try:return original(app,path,read_only=read_only)
    except Exception as exc:primary.append(exc);raise
   return original(app,path,read_only=read_only)
  original_error=c.error_message
  def error(exc):seen.append(exc);return original_error(exc)
  try:
   with patch.object(c,'native_open',side_effect=locked),patch.object(c,'error_message',side_effect=error),patch('akz_stage6.gui.messagebox.showerror',side_effect=lambda title,body,**kwargs:message.append(body)):
    gui.generate_button.configure(state='normal');gui.generate_button.invoke()
   assert primary[0] is seen[0] and gui.checked is None and not gui.outputs
   assert 'FILE_LOCKED:' in message[0] and 'Файл приложения АКЗ занят' in message[0] and 'WinError 32' not in message[0]
   assert win32file.GetFileType(handles[0])==win32con.FILE_TYPE_DISK
   detail={'modal_body':message[0],'primary_same_object_to_error_message':True,'live_lock_after_actual_Tk_button_handler':True,'preview_invalidated':True,'success_not_reported':True}
  finally:
   for h in handles:h.Close()
   root.destroy()
  c.recover_pending();return detail
 check('expanded_launcher_zero_R2_primary_error_preserved_through_Tk_handler',gui_zero_lock)
def gui_preview():
 root=tk.Tk();root.withdraw();gui=Launcher(root)
 try:
  gui.vom.set(str(vom));gui.catalog.set(str(catalog));gui.check()
  assert gui.checked is not None
  text=gui.result.get('1.0','end')
  assert all(f'№ АКЗ={n}: READY' in text for n in (1,2,3)) and '№ АКЗ=0: EMPTY' in text
  assert all(gui.variant_vars[n].get() for n in (1,2,3)) and not gui.variant_vars[0].get()
  assert 'NOT_PROVEN' in text and '37.2' in text and '1.412' in text
  return {'text':text,'ready_variants':[1,2,3],'zero':'EMPTY'}
 finally:root.destroy()
check('actual_Tk_launcher_real_group_preview_materials_totals',gui_preview)
unit=subprocess.run([sys.executable,'-B','-m','unittest','discover','-s','tests','-p','test_stage6*.py','-v'],cwd=ROOT,capture_output=True,text=True,encoding='utf-8')
(OUT/'unit_tests.txt').write_text(unit.stdout+unit.stderr,encoding='utf-8');assert unit.returncode==0,unit.stderr
dump('tested_product_after.json',{str(p):c.sha256(p) for p in product})
assert json.loads((OUT/'tested_product_before.json').read_text())==json.loads((OUT/'tested_product_after.json').read_text())
print('ALL_TARGETED_PASS',len(results),'UNIT_PASS',flush=True)
