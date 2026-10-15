from pathlib import Path
import sys,json,shutil
from unittest.mock import patch
import tkinter as tk
import win32file,win32con
E=Path(__file__).resolve().parent;R=E.parents[1];sys.path.insert(0,str(R))
from akz_stage6 import core as c
from akz_stage6.gui import Launcher
F=E/'baseline_failure/fixtures';O=E/'baseline_failure/output';F.mkdir(parents=True);O.mkdir()
old=json.loads((R/'validation/stage6_akz_part1_20261009_085822/preview.json').read_text('utf-8'))['payload']
cat=R/'DEV/Stage6_Fixtures/stage6_akz_part1_20261009_085822_release3/TEST_Справочник марок.xlsx'
def cp(p,n):q=F/n;shutil.copyfile(p,q);return q
reg=cp(c.REGISTRY,'DEV_registry.xlsx');act=cp(c.ACT,'ACT_template.xlsx');appfile=cp(c.APPLICATION,'APPLICATION_template.xlsx');catalog=cp(cat,'TEST_catalog.xlsx');vom=cp(old['vom']['path'],'selected_VOM.xls')
handles=[];proof={};original=c.native_open

def locked(excel,path,*,read_only):
 if Path(path).name=='ТЕСТ Приложение к АКЗ №0.xlsx' and not read_only:
  proof.update(path=str(path),healthy_sha=c.sha256(path),share=0)
  handles.append(win32file.CreateFile(str(path),win32con.GENERIC_READ,0,None,win32con.OPEN_EXISTING,0,None))
  proof['lstat_while_locked']={'dev':Path(path).lstat().st_dev,'ino':Path(path).lstat().st_ino}
  try:return original(excel,path,read_only=read_only)
  except Exception as exc:proof.update(core_code=getattr(exc,'code',None),core_message=str(exc));raise
 return original(excel,path,read_only=read_only)
with patch.multiple(c,REGISTRY=reg,ACT=act,APPLICATION=appfile,FIXTURES=F,OUTPUT=O,PINS={act:c.sha256(act),appfile:c.sha256(appfile)}):
 p=c.preview('Гараж Казахстан','9/2026',vom,catalog)
 root=tk.Tk();root.withdraw();gui=Launcher(root);gui.checked=p
 try:
  with patch.object(c,'native_open',side_effect=locked),patch('akz_stage6.gui.messagebox.showerror') as msg:gui.generate()
  proof.update(GUI_message=msg.call_args.args[1],lock_held_through_GUI_return=bool(handles),pending_after_GUI=[str(x) for x in O.glob('.pending-*')],partial_published=[str(x) for x in O.iterdir() if not x.name.startswith('.pending-')])
  assert proof['core_code']=='FILE_LOCKED' and proof['GUI_message'].startswith('OTHER_ERROR') and proof['pending_after_GUI'] and not proof['partial_published']
  (E/'baseline_failure_reproduced.json').write_text(json.dumps(proof,ensure_ascii=False,indent=2),encoding='utf-8')
 finally:
  for h in handles:h.Close()
  root.destroy()
# Reproduction-only cleanup, after release. All paths were just created by this harness.
for p in O.glob('.pending-*'):
 assert p.resolve().parent==O.resolve() and O.resolve().is_relative_to(E.resolve())
 for f in p.iterdir():
  assert f.name in ('ТЕСТ АКЗ №0.xlsx','ТЕСТ Приложение к АКЗ №0.xlsx') and f.is_file();f.unlink()
 p.rmdir()
print('BASELINE FAILURE REPRODUCED: FILE_LOCKED masked by cleanup; handle held until GUI returned',flush=True)
