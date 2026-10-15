from common import *
from unittest.mock import patch
from contextlib import ExitStack
import types
from stage9_app.ownership import request_empty_excel_close
record={'pid':17,'image':r'C:\Office\EXCEL.EXE','creation_time':'BIRTH','actual_child':True,'owned':True,'handle_source':'CREATE_PROCESS','launcher_pid':os.getpid(),'quit_returned':True,'workbooks_before_quit':0}
sent=[];cases=[]
with ExitStack() as s:
    for p in [patch('pythoncom._GetInterfaceCount',return_value=0),patch('win32process.GetProcessId',return_value=17),patch('win32process.GetProcessTimes',return_value={'CreationTime':types.SimpleNamespace(isoformat=lambda:'BIRTH')}),patch('stage9_app.ownership.process_image',return_value=record['image']),patch('win32gui.EnumWindows',side_effect=lambda fn,arg:fn(101,arg)),patch('win32process.GetWindowThreadProcessId',return_value=(1,17)),patch('win32gui.GetClassName',return_value='XLMAIN'),patch('win32gui.IsWindow',return_value=True),patch('win32gui.PostMessage',side_effect=lambda *a:sent.append(a))]:s.enter_context(p)
    for name,patcher in [('birth_changed',patch('win32process.GetProcessTimes',return_value={'CreationTime':types.SimpleNamespace(isoformat=lambda:'OTHER')})),('handle_invalid',patch('win32process.GetProcessId',side_effect=OSError('invalid handle'))),('window_changes_owner',patch('win32process.GetWindowThreadProcessId',side_effect=[(1,17),(1,99)])),('window_changes_class',patch('win32gui.GetClassName',side_effect=['XLMAIN','OTHER'])),('window_destroyed',patch('win32gui.IsWindow',return_value=False))]:
        sent.clear();error=None
        with patcher:
            try:request_empty_excel_close(dict(record),77,lambda *a:None)
            except (RuntimeError,OSError) as exc:error=repr(exc)
        assert not sent,name;cases.append({'case':name,'no_WM_CLOSE':True,'error':error})
    for name,delta in [('not_actual_child',{'actual_child':False}),('not_original_source',{'handle_source':'OPEN_PROCESS'})]:
        sent.clear()
        try:request_empty_excel_close(record|delta,77,lambda *a:None)
        except RuntimeError:pass
        else:raise AssertionError(name)
        assert not sent;cases.append({'case':name,'no_WM_CLOSE':True})
save(OUT/'WINDOW_CLOSE_EXTENDED_SOURCE.json',{'status':'PASS','source_only':True,'source_sha256':sha(ROOT/'stage9_app/ownership.py'),'cases':cases})
print('Extended ownership races and invalid handles',len(cases),'PASS',flush=True)
