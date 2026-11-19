"""Source-only ownership regression; never sends input to a real window."""
from common import *
from unittest.mock import patch
import types
from stage9_app.ownership import request_empty_excel_close
record={'pid':123,'creation_time':'BIRTH','image':r'C:\Office\EXCEL.EXE','actual_child':True,
        'owned':True,'handle_source':'CREATE_PROCESS','launcher_pid':os.getpid(),
        'quit_returned':True,'workbooks_before_quit':0}
calls=[];events=[];cases=[]
def enumeration(fn,arg):
    for hwnd in (101,102):fn(hwnd,arg)
patches=[patch('pythoncom._GetInterfaceCount',return_value=0),patch('win32process.GetProcessId',return_value=123),
    patch('win32process.GetProcessTimes',return_value={'CreationTime':types.SimpleNamespace(isoformat=lambda:'BIRTH')}),
    patch('stage9_app.ownership.process_image',return_value=record['image']),
    patch('win32gui.EnumWindows',side_effect=enumeration),patch('win32process.GetWindowThreadProcessId',side_effect=lambda h:(1,123 if h==101 else 999)),
    patch('win32gui.GetClassName',return_value='XLMAIN'),patch('win32gui.IsWindow',return_value=True),
    patch('win32gui.PostMessage',side_effect=lambda *a:calls.append(a))]
from contextlib import ExitStack
with ExitStack() as stack:
    for p in patches:stack.enter_context(p)
    request_empty_excel_close(dict(record),77,lambda k,v:events.append((k,v)))
    assert calls==[(101,0x10,0,0)] and events[0][1]['handle_validated']
    cases.append('only_proven_owned_window_receives_WM_CLOSE')
    for label,delta,expected_error in [('foreign_launcher',{'launcher_pid':-99},True),('not_owned',{'owned':False},True),
        ('missing_quit',{'quit_returned':False},False),('books_remain',{'workbooks_before_quit':1},False)]:
        calls.clear()
        try:request_empty_excel_close(record|delta,77,lambda *a:None)
        except RuntimeError:assert expected_error
        else:assert not expected_error
        assert not calls;cases.append(label)
    for label,target in [('reused_pid','win32process.GetProcessId'),('changed_image','stage9_app.ownership.process_image'),('live_com','pythoncom._GetInterfaceCount')]:
        calls.clear()
        with patch(target,return_value=999 if label=='reused_pid' else 'OTHER.EXE' if label=='changed_image' else 1):
            try:request_empty_excel_close(dict(record),77,lambda *a:None)
            except RuntimeError:pass
            else:raise AssertionError(label)
        assert not calls;cases.append(label)
save(OUT/'WINDOW_CLOSE_GUARD_SOURCE.json',{'status':'PASS','source_only':True,'source_sha256':sha(ROOT/'stage9_app/ownership.py'),'cases':cases})
print('Owned empty-window close guard',len(cases),'PASS')
