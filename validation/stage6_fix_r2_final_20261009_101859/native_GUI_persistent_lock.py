from common import *
import tkinter as tk
import win32file,win32con
from tkinter import messagebox
F,O=setup('native_GUI');old=snapshot(O/'OLD_GOOD');unrelated=sha(O/'unrelated.xlsx')
product=load(E/'final_tests_product_BEFORE.json');assert {p:sha(p) for p in product}==product
handles=[];proof=[];primary=[];original=c.native_open;show=messagebox.showerror

def locked(excel,path,*,read_only):
 if Path(path).name==APPNAME and not read_only:
  healthy=sha(path);assert healthy==sha(F/'APPLICATION_template.xlsx')
  handles.append(win32file.CreateFile(str(path),win32con.GENERIC_READ,0,None,win32con.OPEN_EXISTING,0,None))
  try:return original(excel,path,read_only=read_only)
  except Exception as exc:primary.append(exc);proof.append({'path':str(path),'healthy_sha':healthy,'share':0,'code':getattr(exc,'code',None),'winerror':getattr(exc.__cause__,'winerror',None)});raise
 return original(excel,path,read_only=read_only)

with isolated(F,O):
 root=tk.Tk();gui=Launcher(root);gui.vom.set(str(F/'selected_VOM.xls'));gui.catalog.set(str(F/'TEST_catalog.xlsx'));gui.check();p=gui.checked
 assert p is not None and str(gui.generate_button['state'])=='normal'
 def shown(title,message,**kw):
  assert handles and all(win32file.GetFileType(h)==win32con.FILE_TYPE_DISK for h in handles)
  assert primary[0].code=='FILE_LOCKED' and primary[0].cleanup_error.code=='FILE_LOCKED'
  assert message.startswith('FILE_LOCKED') and 'Закройте файл в Excel' in message and proof[0]['path'] in message
  save('native_GUI_message_while_locked.json',{'title':title,'message':message,'proof':proof,'cleanup_error':str(primary[0].cleanup_error),'handle_live_during_modal':True,'pending':[str(x) for x in O.glob('.pending-*')],'ownership_records':[load(x) for x in O.glob('.stage6-pending-*.json')]})
  print('MODAL_READY_LOCK_STILL_HELD',flush=True)
  answer=show(title,message,**kw);root.after(100,root.quit);return answer
 print('READY_CLICK_GENERATE',flush=True)
 try:
  with patch.object(c,'native_open',side_effect=locked),patch('akz_stage6.gui.messagebox.showerror',side_effect=shown):root.mainloop()
  assert handles and all(win32file.GetFileType(h)==win32con.FILE_TYPE_DISK for h in handles)
  pending=list(O.glob('.pending-*'));assert len(pending)==1 and gui.checked is None
  save('native_GUI_after_return_lock_still_held.json',{'handle_live_after_GUI_return':True,'preview_invalidated':True,'pending':str(pending[0]),'pending_files':[str(x) for x in pending[0].iterdir()],'no_partial_published':set(x.name for x in O.iterdir() if not x.name.startswith('.'))=={'OLD_GOOD','unrelated.xlsx'},'old_good_unchanged':snapshot(O/'OLD_GOOD')==old})
 finally:
  for h in handles:h.Close()
  root.destroy()
 # No harness cleanup: normal product generation must recover its own pending.
 target=c.generate(p,'RETRY_AFTER_RELEASE')
 assert load(target/'PAIR.json')['state']=='COMPLETE_TEST_PAIR'
 assert not list(O.glob('.pending-*')) and not list(O.glob('.stage6-pending-*.json'))
 assert snapshot(O/'OLD_GOOD')==old and sha(O/'unrelated.xlsx')==unrelated
 save('native_GUI_release_cleanup_retry.json',{'pending_removed_by_product':True,'ownership_records_removed':True,'complete_pair':str(target),'sha256':snapshot(target),'unrelated_unchanged':True,'old_good_unchanged':True,'audit':AUDIT})
 assert {p:sha(p) for p in product}==product
 print('NATIVE_GUI_RELEASE_CLEANUP_RETRY_PASS',flush=True)
