"""Visible unmodified DEV App; passive widget snapshots, all input via Computer Use."""
import sys,json,os,tkinter as tk
from pathlib import Path
RUN=Path(__file__).resolve().parent;sys.path.insert(0,str(RUN.parents[1]));sys.dont_write_bytecode=True
from stage8_closure import office
from stage8_closure.gui import App
def save(p,v):p.write_bytes(json.dumps(v,ensure_ascii=False,indent=2).encode('utf8'))
with office.diagnostic_scope(RUN/'gui_office'):
    window=tk.Tk();app=App(window)
    save(RUN/'GUI_PROCESS.json',{'pid':os.getpid(),'visible':True,'product':'stage8_closure.gui.App','no_input_patches':True})
    def observe():
        save(RUN/'GUI_CURRENT.json',{'text':app.text.get('1.0','end'),'close':str(app.close_button['state']),'restore':str(app.restore_button['state']),'confirmation':app.confirm.get(),'sandbox':str(app.sandbox),'plan':app.plan,'restore_plan':app.restore_plan})
        window.after(300,observe)
    observe();window.mainloop()
