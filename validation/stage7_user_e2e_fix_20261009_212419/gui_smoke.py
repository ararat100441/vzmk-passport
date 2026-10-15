"""Read-only GUI preview of retained status/scope; never call a generator."""
from pathlib import Path
import sys
from unittest.mock import patch
import tkinter as tk
sys.dont_write_bytecode=True
HERE=Path(__file__).resolve().parent;sys.path.insert(0,str(HERE.parents[1]))
from stage7_quality.gui import launch
from stage7_quality.provenance import read,write

manifest=read(HERE/'PREFLIGHT.json')['manifest']['path']
actual_tk=tk.Tk;captured={};errors=[]
def hidden_root():
    root=actual_tk();root.withdraw()
    def drive():
        children=root.winfo_children()
        buttons={w.cget('text'):w for w in children if isinstance(w,tk.Button)}
        buttons['Выбрать manifest и проверить'].invoke()
        content=[w for w in children if isinstance(w,tk.Text)][0].get('1.0','end')
        assert '9/2026-1' in content and 'SAMPLE_ONLY' in content and '3 назначенные марки / 3 изделия' in content
        assert 'Официальная выдача запрещена' in content
        buttons['Сформировать TEST DOCX'].invoke()
        assert errors and 'Подтвердите' in errors[0]
        captured['preview']=content;root.destroy()
    root.mainloop=drive
    return root
with patch('tkinter.Tk',hidden_root),patch('tkinter.filedialog.askopenfilename',return_value=manifest),patch('tkinter.messagebox.showerror',side_effect=lambda title,content:errors.append(content)),patch('stage7_quality.gui.generate_quality',side_effect=AssertionError('No generation without confirmation')):
    launch()
write(HERE/'GUI_SMOKE.json',{'status':'PASS','mode':'Hidden Tk callback smoke','preview':captured['preview'],
    'generator_calls':0,'confirmation_required':True,'live_user_ui_acceptance':'NOT_RUN'})
print('DEV interface status/scope retained; no generator called')
