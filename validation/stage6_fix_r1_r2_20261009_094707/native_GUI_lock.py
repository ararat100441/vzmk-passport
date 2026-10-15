from pathlib import Path
import sys,json
from unittest.mock import patch
import tkinter as tk
import win32file,win32con
E=Path(__file__).resolve().parent;sys.path.insert(0,str(E.parents[1]))
from akz_stage6 import core as c
from akz_stage6 import gui as g
old=json.loads((E.parents[1]/'validation/stage6_akz_part1_20261009_085822/preview.json').read_text('utf-8'))['payload']
cat=E.parents[1]/'DEV/Stage6_Fixtures/stage6_akz_part1_20261009_085822_release3/TEST_Справочник марок.xlsx'
p=c.preview('Гараж Казахстан','9/2026',old['vom']['path'],cat)
root=tk.Tk();launcher=g.Launcher(root)
launcher.vom.set(old['vom']['path']);launcher.catalog.set(str(cat));launcher.checked=p
launcher.result.insert('end','Native GUI lock validation: real Windows CreateFile share=0 on valid pending application.\n')
original=c.native_open;show=g.messagebox.showerror;proof={}
def locked(app,path,*,read_only):
 if Path(path).name=='ТЕСТ Приложение к АКЗ №0.xlsx' and not read_only:
  proof.update(path=str(path),valid_sha256=c.sha256(path),share=0)
  assert c.sha256(path)==c.sha256(c.APPLICATION)
  h=win32file.CreateFile(str(path),win32con.GENERIC_READ,0,None,win32con.OPEN_EXISTING,0,None)
  try:return original(app,path,read_only=read_only)
  finally:h.Close()
 return original(app,path,read_only=read_only)
def actual_dialog(title,message,**kw):
 proof.update(title=title,message=message)
 (E/'native_GUI_lock_on_show.json').write_text(json.dumps(proof,ensure_ascii=False,indent=2),encoding='utf-8')
 return show(title,message,**kw)
def run():
 with patch.object(c,'OUTPUT',E/'native_GUI_output'),patch.object(c,'native_open',side_effect=locked),patch.object(g.messagebox,'showerror',side_effect=actual_dialog):launcher.generate()
 proof.update(dialog_returned=True,output_empty=not list((E/'native_GUI_output').iterdir()),preview_invalidated=launcher.checked is None)
 assert proof['output_empty'] and proof['preview_invalidated'] and 'FILE_LOCKED' in proof['message']
 (E/'native_GUI_lock_result.json').write_text(json.dumps(proof,ensure_ascii=False,indent=2),encoding='utf-8')
 root.after(100,root.destroy)
root.after(400,run)
root.mainloop()
