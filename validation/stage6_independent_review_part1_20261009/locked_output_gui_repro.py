"""Targeted reproduction: valid pending XLSX locked by another file handle."""
from review import *
from akz_stage6 import core as c
from akz_stage6.gui import Launcher
import tkinter as tk
import win32file,win32con
F=E/'sandbox/fixtures';O=E/'sandbox/output'
act=F/'ACT_template.xlsx';application=F/'APPLICATION_template.xlsx'
original_open=c.native_open
lock_proof=[]
def locked_open(app,path,*,read_only):
    if Path(path).name=='ТЕСТ Приложение к АКЗ №0.xlsx' and not read_only:
        assert sha(path)==sha(application),'COPY MUST BE VALID EXACT TEMPLATE'
        handle=win32file.CreateFile(str(path),win32con.GENERIC_READ,0,None,win32con.OPEN_EXISTING,0,None)
        lock_proof.append({'path':str(path),'valid_template_sha256':sha(application),'sharing_mode':0,'handle_opened':True})
        try:return original_open(app,path,read_only=read_only)
        finally:handle.Close()
    return original_open(app,path,read_only=read_only)
with patch.multiple(c,REGISTRY=F/'DEV_registry.xlsx',ACT=act,APPLICATION=application,FIXTURES=F,OUTPUT=O,
                    PINS={act:sha(act),application:sha(application)}):
    p=c.preview('Гараж Казахстан','9/2026',F/'SYNTHETIC_17_rows.xlsx',F/'TEST_catalog.xlsx',synthetic=True)
    root=tk.Tk();root.withdraw()
    try:
        gui=Launcher(root);gui.checked=p
        before={x.name for x in O.iterdir()}
        with patch.object(c,'native_open',side_effect=locked_open),patch('akz_stage6.gui.messagebox.showerror') as error:
            gui.generate()
        assert error.call_count==1 and lock_proof
        message=error.call_args.args[1]
        after={x.name for x in O.iterdir()}
        assert before==after and not list(O.glob('.pending-*'))
        assert 'формат или расширение' in message and 'занят' not in message.lower()
        save('locked_output_gui_reproduction.json',{'lock_proof':lock_proof,'displayed_title':error.call_args.args[0],
            'displayed_message':message,'safe_rollback':True,'published_pair':False,
            'acceptance_actionable_lock_message':'FAIL','finding':'S6-P1-R2'})
        print('REPRODUCED S6-P1-R2: valid locked file misreported as invalid format; rollback PASS')
    finally:root.destroy()
t=load(E/'independent_tests.json')
for test in t['results']:
    if test['name']=='locked_second_output_actionable_native_error':
        test['name']='locked_second_output_safe_native_error'
        test['assessment']='Safe STOP/rollback PASS; misleading lock message is S6-P1-R2, not a clarity PASS'
save('independent_tests.json',t)
