from pathlib import Path
import sys,json
ROOT=Path(__file__).resolve().parents[2];OUT=Path(__file__).resolve().parent
sys.path.insert(0,str(ROOT))
from akz_stage6 import core as c
from akz_stage6.gui import Launcher
import tkinter as tk
events=[]
def capture(name,original):
 def wrapped(self,*args):
  value=original(self,*args)
  events.append({'action':name,'preview':None if self.checked is None else {n:g['status'] for n,g in self.checked.payload['groups'].items()},
                 'outputs':[str(x) for x in self.outputs],'text':self.result.get('1.0','end'),
                 'excel_books':[{'name':b.Name,'read_only':bool(b.ReadOnly)} for app in self.open_apps for b in app.Workbooks]})
  (OUT/'native_GUI_events.json').write_text(json.dumps(events,ensure_ascii=False,indent=2),encoding='utf-8')
  return value
 return wrapped
for method in ('check','generate','open'):setattr(Launcher,method,capture(method,getattr(Launcher,method)))
root=tk.Tk();gui=Launcher(root)
gui.vom.set(str(c.PROJECT/'Паспорт/9 от 06.10.2026/ВОМ отгрузочный счет 183  (5).xls'))
gui.catalog.set(str(c.FIXTURES/'stage6_part2_20261009_132204/TEST_Справочник марок.xlsx'))
def close():
 for app in gui.open_apps:
  for book in list(app.Workbooks):book.Close(SaveChanges=False)
  app.Quit()
 root.destroy()
root.protocol('WM_DELETE_WINDOW',close)
print('NATIVE_GUI_READY',flush=True)
root.mainloop()
print('NATIVE_GUI_CLOSED',flush=True)
