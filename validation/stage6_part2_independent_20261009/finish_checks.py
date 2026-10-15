from review import *
from zipfile import ZipFile
from openpyxl import load_workbook
from akz_stage6 import core as c,part2 as p2
from akz_stage6.gui import Launcher
from unittest.mock import patch
import tkinter as tk
import win32file,win32con
m=json.loads((IMPL/'REVIEW_MANIFEST.json').read_text(encoding='utf-8'))
proof={'receipts':[],'package_checks':[],'old_zero':[]}
preview=json.loads((OUT/'independent_preview.json').read_text(encoding='utf-8'))
for x in m['six_outputs']:
    path=Path(x['path']);assert sha(path)==x['sha256']
    with ZipFile(path) as z:
        forbidden=[n for n in z.namelist() if n.startswith(('xl/externalLinks/','xl/media/','xl/activeX/','xl/embeddings/')) or n.endswith('vbaProject.bin')]
        assert not forbidden,forbidden
        for name in z.namelist():
            if name.startswith('xl/') and name.endswith('.xml'):
                data=z.read(name).decode('utf-8');assert 'RANDBETWEEN' not in data.upper() and 'RAND(' not in data.upper() and '#REF!' not in data
        proof['package_checks'].append({'path':str(path),'no_externalLinks_media_activex_embedded_vba':True,'no_random_or_REF_any_xl_xml':True})
    if x['kind']=='act':
        receipt=json.loads((path.parent/'PAIR.json').read_text(encoding='utf-8'));n=x['variant'];g=preview['groups'][str(n)]
        assert receipt['act_id']==g['seal'][:16] and receipt['official']=='STOP' and receipt['state']=='COMPLETE_TEST_PAIR' and receipt['variant']==n
        for name,h in receipt['sha256'].items():assert sha(path.parent/name)==h
        proof['receipts'].append({'variant':n,'path':str(path.parent/'PAIR.json'),'sha256':sha(path.parent/'PAIR.json'),'independent_preview_seal':g['seal'],'act_id_matches':True})
old=ROOT/'validation/stage6_fix_r1_r2_20261009_094707/NEW_TEST_9_2026'
for name,h in [('ТЕСТ АКЗ №0.xlsx','b640796570a8d04b6bca404205d99eae4c7005c700587812ed25bc99e715eb1a'),('ТЕСТ Приложение к АКЗ №0.xlsx','7766c92df82df2484eb82cd6917769f5d9ede912b3f0e8f26e3f8aaabdd357f2'),('PAIR.json','eaad2385ef401a21f9203db9648530f71939ec4933296f309adb7db19bedbb8f')]:
    assert sha(old/name)==h;proof['old_zero'].append({'path':str(old/name),'sha256':h,'unchanged':True})
box=OUT/'gui_isolated';box.mkdir(exist_ok=False)
vom=c.PROJECT/'Паспорт/9 от 06.10.2026/ВОМ отгрузочный счет 183  (5).xls'
root=tk.Tk();root.withdraw();gui=Launcher(root)
try:
    gui.vom.set(str(vom));gui.catalog.set(m['fixture']['path']);gui.check();assert gui.checked is not None
    text=gui.result.get('1.0','end');assert all(f'№ АКЗ={n}: READY' in text for n in (1,2,3)) and '№ АКЗ=0: EMPTY' in text
    assert '∑1+2=180' in text and 'NOT_PROVEN' in text and all(gui.variant_vars[n].get() for n in (1,2,3))
    proof['actual_Tk_preview']={'text':text,'ready':[1,2,3],'zero':'EMPTY','no_derived_80_for_variant2':True}
    old_fixture=c.FIXTURES/'stage6_akz_part1_20261009_085822_release3/TEST_Справочник марок.xlsx'
    batch=p2.preview_groups('Гараж Казахстан','9/2026',vom,old_fixture);gui.checked=batch
    for n,var in gui.variant_vars.items():var.set(n==0)
    handles=[];primary=[];seen=[];messages=[];opened=c.native_open;error=c.error_message
    def lock(app,path,*,read_only):
        if Path(path).name=='ТЕСТ Приложение к АКЗ №0.xlsx' and not read_only:
            handles.append(win32file.CreateFile(str(path),win32con.GENERIC_READ,0,None,win32con.OPEN_EXISTING,0,None))
            try:return opened(app,path,read_only=read_only)
            except Exception as exc:primary.append(exc);raise
        return opened(app,path,read_only=read_only)
    def message(exc):seen.append(exc);return error(exc)
    try:
        with patch.object(c,'OUTPUT',box),patch.object(c,'native_open',side_effect=lock),patch.object(c,'error_message',side_effect=message),patch('akz_stage6.gui.messagebox.showerror',side_effect=lambda title,body,**kwargs:messages.append(body)):
            gui.generate_button.configure(state='normal');gui.generate_button.invoke()
        assert primary[0] is seen[0] and 'FILE_LOCKED:' in messages[0] and 'Файл приложения АКЗ занят' in messages[0]
        assert not gui.outputs and gui.checked is None and win32file.GetFileType(handles[0])==win32con.FILE_TYPE_DISK
        proof['actual_Tk_R2_handler']={'message':messages[0],'primary_object_preserved':True,'no_success':True,'live_handle_through_handler':True,'preview_invalidated':True}
    finally:
        for h in handles:h.Close()
    with patch.object(c,'OUTPUT',box):c.recover_pending()
finally:root.destroy()
write('additional_checks.json',proof)
print('PACKAGES RECEIPTS OLD_ZERO GUI_PREVIEW GUI_R2 PASS',flush=True)
