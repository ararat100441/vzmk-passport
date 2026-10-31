"""Automated Tk component test of byte-identical sources shipped with RC1.

No external Windows input, screenshots, or Computer Use. Chooser callbacks are
stubbed to select only acceptance-copy paths. Actual EXE smoke is run separately.
"""
from pathlib import Path
import hashlib,json,os,sys,tempfile
from unittest.mock import patch

def main():
    import argparse
    parser=argparse.ArgumentParser();parser.add_argument('--exe',required=True);parser.add_argument('--base',required=True);parser.add_argument('--output',required=True);args=parser.parse_args()
    exe=Path(args.exe).resolve();base=Path(args.base);out=Path(args.output).resolve();out.mkdir(exist_ok=False)
    sys.path.insert(0,str(exe.parent/'_internal'))
    os.environ['LOCALAPPDATA']=str(out/'local')
    from production_app.gui import App
    from production_app.storage import local_root,read
    import tkinter as tk
    cfg=json.loads((base/'RESULT.json').read_text(encoding='utf8'))['cfg']
    expected={k:cfg[k] for k in ('registry','documents','output')}
    root=tk.Tk();root.withdraw();app=App(root)
    for callback in root.tk.call('after','info'):root.after_cancel(callback)
    assert not app.cfg.get('registry') and not (local_root()/'settings.json').exists()
    errors=[]
    def buttons(parent):
        result=[]
        for w in parent.winfo_children():
            if w.winfo_class()=='TButton':result.append(w)
            result.extend(buttons(w))
        return result
    with patch('production_app.gui.filedialog.askopenfilename',return_value=expected['registry']),patch('production_app.gui.filedialog.askdirectory',side_effect=[expected['documents'],expected['output']]),patch('production_app.gui.messagebox.showerror',side_effect=lambda *a:errors.append(a)):
        app.settings();dialog=next(w for w in root.winfo_children() if isinstance(w,tk.Toplevel))
        choose=[w for w in buttons(dialog) if w.cget('text')=='Выбрать…']
        assert len(choose)==3
        for b in choose:b.invoke()
        next(w for w in buttons(dialog) if w.cget('text')=='Сохранить выбранные рабочие пути').invoke()
    assert not errors,errors
    stored=read(local_root()/'settings.json');assert all(stored[k]==v for k,v in expected.items())
    root.destroy()
    root=tk.Tk();root.withdraw();again=App(root)
    for callback in root.tk.call('after','info'):root.after_cancel(callback)
    assert all(again.cfg[k]==v for k,v in expected.items());root.destroy()
    data={'status':'PASS','scope':'Tk component integration using exact shipped source files; not external GUI automation',
        'settings_paths':expected,'persisted_between_instances':True,'silent_dev_default':False,
        'gui_source_sha256':hashlib.sha256((exe.parent/'_internal/production_app/gui.py').read_bytes()).hexdigest()}
    (out/'RESULT.json').write_text(json.dumps(data,ensure_ascii=False,indent=2),encoding='utf8')
    print(json.dumps(data,ensure_ascii=False,indent=2))

if __name__=='__main__':main()
