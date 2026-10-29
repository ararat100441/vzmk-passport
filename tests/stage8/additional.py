"""Stage 8 targeted boundary and GUI tests; does not rerun older stages."""
import json
from pathlib import Path
import sys
import traceback
from unittest.mock import patch
import tkinter as tk
import win32file
import win32con
from stage8_closure import core,storage as S
from stage8_closure.fixtures import create
from stage8_closure.gui import App

RUN=Path(sys.argv[1]).resolve();OUT=RUN/'additional';OUT.mkdir(exist_ok=False)
before=core.code_bindings();records=[]


def fixture():
    r=create(RUN/'sandboxes');return Path(r['sandbox']),Path(r['request'])


def overwrite(p,data):p.write_text(json.dumps(data,ensure_ascii=False,indent=2),encoding='utf8')


def expect(fn,code):
    try:fn()
    except Exception as exc:
        assert code in str(exc),(code,str(exc));return str(exc)
    raise AssertionError('Expected STOP: '+code)


def case(name,fn):
    print(name,'START',flush=True)
    try:r={'name':name,'status':'PASS','detail':fn()}
    except Exception as exc:r={'name':name,'status':'FAIL','error':str(exc),'traceback':traceback.format_exc()}
    S.write(OUT/(name+'.json'),r);records.append(r);print(name,r['status'],flush=True)


def write_guards():
    root,q=fixture();p=core.preview(root,q);attempts=[]
    def fault(point,target):
        if point not in ('before_replace','after_replace'):return
        for path in [Path(p['payload']['registry']),root/'inputs/CONTROL.json']:
            try:
                handle=win32file.CreateFile(str(path),win32con.GENERIC_WRITE,
                    win32con.FILE_SHARE_READ|win32con.FILE_SHARE_WRITE|win32con.FILE_SHARE_DELETE,
                    None,win32con.OPEN_EXISTING,0,None)
            except Exception as exc:
                assert exc.winerror==32;attempts.append({'point':point,'path':str(path),'winerror':32})
            else:handle.Close();raise AssertionError('External writer could open guarded file')
    r=core.close(p,core.confirmation(p),_fault=fault)
    assert r['status']=='CLOSED_TEST' and len(attempts)==4
    return {'sandbox':str(root),'attempts':attempts}
case('deny_external_writers_before_and_after_publish',write_guards)


def duplicate_row():
    from stage8_closure.office import excel_session
    root,q=fixture();reg=root/'registry/УЗК_TEST.xlsm'
    with excel_session() as excel:
        b=excel.Workbooks.Open(Filename=str(reg),UpdateLinks=0,ReadOnly=False,AddToMru=False)
        try:
            s=b.Worksheets('УЗК');s.Cells(27,1).Value2='STAGE8 ИСКУССТВЕННЫЙ ПРОЕКТ';s.Cells(27,2).Value2='S8-A';s.Cells(27,3).Value2=12;b.Save()
        finally:b.Close(SaveChanges=False)
    # Fresh fixture seed after deliberate fixture-only construction, prior to any closure.
    m=S.read(root/'SANDBOX.json');m['initial_registry_sha256']=S.sha(reg);m.pop('seal');m['seal']=S.digest(m);overwrite(root/'SANDBOX.json',m)
    return {'sandbox':str(root),'stop':expect(lambda:core.preview(root,q),'UZK_MARK_MISSING_OR_DUPLICATE')}
case('duplicate_target_registry_rows_stop',duplicate_row)


def membership():
    root,q=fixture();d=S.read(root/'inputs/VIK.json');d['shipment_id']='OTHER-SHIPMENT';overwrite(root/'inputs/VIK.json',d)
    request=S.read(q)
    for entry in request['documents']:
        if entry['kind']=='VIK':entry['sha256']=S.sha(entry['path'])
    overwrite(q,request)
    return {'sandbox':str(root),'stop':expect(lambda:core.preview(root,q),'EVIDENCE_MEMBERSHIP_MISMATCH')}
case('document_from_other_shipment_stop',membership)


def quantity_mismatch():
    root,q=fixture();file=root/'inputs/CONTROL.json';d=S.read(file);d['uzk'][0]['quantity']=2;overwrite(file,d)
    request=S.read(q);request['control']['sha256']=S.sha(file);overwrite(q,request)
    return {'sandbox':str(root),'stop':expect(lambda:core.preview(root,q),'UZK_CONTROL_INCOMPLETE')}
case('partially_confirmed_quantity_stop',quantity_mismatch)


def forged_preview():
    root,q=fixture();p=core.preview(root,q);p['payload']['deductions'][0]['after']=999
    e=expect(lambda:core.close(p,'anything'),'PREVIEW_TAMPERED');assert not S.history(root)['events']
    return {'sandbox':str(root),'stop':e}
case('preview_edit_does_not_authorize_write',forged_preview)


def gui():
    root,q=fixture();window=tk.Tk();window.withdraw();app=App(window);messages=[];confirmations=[]
    try:
        app.load(q);window.update_idletasks()
        text=app.text.get('1.0','end')
        for phrase in ('Проект:','Паспорт:','Выбранная отгрузка:','Состав документов:','Обязательные документы:',
                       'Статус проверки:','Количество подтверждённых изделий УЗК:','S8-A: 12 − 3 = 9','S8-B: 10 − 4 = 6'):
            assert phrase in text,phrase
        assert str(app.close_button['state'])=='normal' and not app.confirm.get()
        with patch('stage8_closure.gui.messagebox.showerror',side_effect=lambda *x:messages.append(x)), \
             patch('stage8_closure.gui.messagebox.askyesno',side_effect=lambda *x,**k:confirmations.append((x,k)) or False):
            app.close_button.invoke();assert messages and not confirmations and not S.history(root)['events']
            app.load(q);app.confirm.set(True);app.close_button.invoke()
            assert confirmations and confirmations[0][1]['default']=='no' and not S.history(root)['events']
        with patch('stage8_closure.gui.messagebox.askyesno',return_value=True):
            app.close_button.invoke()
        assert 'TEST закрыт' in app.text.get('1.0','end') and str(app.close_button['state'])=='disabled'
        app.load(q);assert app.plan['payload']['status']=='ALREADY_CLOSED_TEST' and str(app.close_button['state'])=='disabled'
        # Current Stage 7 selection is inspected without generators or registry writes.
        files=sorted((S.ROOT/'validation').glob('stage7_*/**/STAGE7_SELECTED_TEST_MANIFEST.json'))
        assert files;app.load(files[-1]);blocked=app.text.get('1.0','end')
        assert '9/2026' in blocked and 'SAMPLE_ONLY' in blocked and str(app.close_button['state'])=='disabled'
        S.write(OUT/'GUI_TEXT.json',{'ready_preview':text,'stage7_blocked':blocked,'modal_confirmation_called':True,
            'cancel_writes':0,'close_handler':'PASS','repeat_disabled':True,'manual_user_e2e':'NOT_RUN'})
        return {'sandbox':str(root),'checks':'ready/checkbox/modal cancel/modal accept/reopen/Stage7 block','manual_user_e2e':'NOT_RUN'}
    finally:window.destroy()
case('gui_preview_confirmation_cancel_close_and_block',gui)


def journal_metadata():
    root,q=fixture();crashplan=core.preview(root,q)
    def fault(point,target):
        if point=='after_save':raise OSError('TEST_PENDING')
    expect(lambda:core.close(crashplan,core.confirmation(crashplan),_fault=fault),'TEST_PENDING')
    op=S.history(root)['active'];f=Path(op['preservation']['path']);f.write_bytes(f.read_bytes()+b' ')
    return {'sandbox':str(root),'stop':expect(lambda:core.recover(root),'PRESERVATION_EVIDENCE_DAMAGED')}
case('recovery_evidence_changed_stop',journal_metadata)


after=core.code_bindings();assert before==after
S.write(OUT/'RESULTS.json',{'cases':records,'passed':sum(r['status']=='PASS' for r in records),
    'failed':sum(r['status']=='FAIL' for r in records),'product_before':before,'product_after':after})
sys.exit(1 if any(r['status']=='FAIL' for r in records) else 0)
