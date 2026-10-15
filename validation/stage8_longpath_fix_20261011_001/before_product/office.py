"""Excel ownership is local to Stage 8; keep COM apartment alive per thread.

Office proxy objects may outlive a Python context's local Excel variable.
Repeated CoUninitialize while such proxies exist can invalidate type-info
objects. Initialize once per calling thread; the OS releases it on process exit.
Every owned Excel instance is still explicitly closed and Quit in finally.
"""
from contextlib import contextmanager
from pathlib import Path
import sys
import tempfile
import threading
from uuid import uuid4
from .storage import Stop, now, write

_apartment=threading.local()


@contextmanager
def diagnostic_scope(directory):
    previous = getattr(_apartment,'diagnostics',None)
    _apartment.diagnostics = Path(directory)
    try:
        yield
    finally:
        _apartment.diagnostics = previous


def _error(exc):
    hresult = getattr(exc,'hresult',None)
    return {'type':type(exc).__name__,'message':str(exc),'hresult':hresult,
            'hresult_hex':hex(hresult & 0xffffffff) if hresult is not None else None,
            'notes':getattr(exc,'__notes__',[]),
            'cause':_error(exc.__cause__) if exc.__cause__ is not None else None}


def record_cleanup_error(step, exc):
    record = getattr(_apartment,'current_record',None)
    if record is not None:
        record['cleanup'].append({'step':step,'status':'ERROR',**_error(exc)})


@contextmanager
def excel_session(*, diagnostics=None, workbooks=()):
    import pythoncom
    import win32com.client
    import win32process, win32api, win32con
    if not getattr(_apartment,'initialized',False):
        pythoncom.CoInitialize()
        _apartment.initialized=True
    directory = Path(diagnostics or getattr(_apartment,'diagnostics',None)
                     or Path(tempfile.gettempdir())/'Pasport-stage8-office')/uuid4().hex
    directory.mkdir(parents=True,exist_ok=False)
    record = {'schema':'STAGE8_OWNED_EXCEL_V1','started':now(),
              'expected_workbooks':[str(Path(p).resolve()) for p in workbooks],
              'cleanup':[], 'owned':False}
    excel=None
    existing = set(win32process.EnumProcesses())
    previous_record = getattr(_apartment,'current_record',None)
    _apartment.current_record = record
    try:
        excel=win32com.client.DispatchEx('Excel.Application')
        pid = win32process.GetWindowThreadProcessId(excel.Hwnd)[1]
        process = win32api.OpenProcess(win32con.PROCESS_QUERY_INFORMATION,False,pid)
        try:
            created = win32process.GetProcessTimes(process)['CreationTime'].isoformat()
        finally:
            process.Close()
        record.update(pid=pid,creation_time=created,owned=pid not in existing)
        write(directory/'START.json',record)
        if not record['owned']:
            raise Stop('EXCEL_OWNERSHIP_NOT_PROVEN_NO_QUIT')
        excel.Visible=False
        excel.DisplayAlerts=False
        excel.EnableEvents=False
        excel.AskToUpdateLinks=False
        excel.AutomationSecurity=3
        excel.ScreenUpdating=False
        yield excel
    except Stop:
        raise
    except Exception as exc:
        raise Stop('MICROSOFT_EXCEL_FAILURE: '+str(exc)) from exc
    finally:
        original = sys.exc_info()[1]
        if excel is not None and record['owned']:
            # Independent attempts: a disconnected Workbook must never skip Quit.
            for name, action in (
                    ('CutCopyMode',lambda:setattr(excel,'CutCopyMode',False)),
                    ('Workbooks.Close',lambda:excel.Workbooks.Close()),
                    ('Quit',lambda:excel.Quit())):
                try:
                    action()
                    record['cleanup'].append({'step':name,'status':'PASS'})
                except Exception as exc:
                    record['cleanup'].append({'step':name,'status':'ERROR',**_error(exc)})
        record.update(finished=now(),original_exception=_error(original) if original else None)
        failures = [r for r in record['cleanup'] if r['status']=='ERROR']
        # Do not retain an application proxy in the cleanup exception traceback.
        action=None
        excel=None
        _apartment.current_record = previous_record
        try:
            write(directory/'END.json',record)
        except Exception as exc:
            if original is not None:
                original.add_note('EXCEL_DIAGNOSTIC_WRITE_FAILED: '+str(exc))
            else:
                raise Stop('EXCEL_DIAGNOSTIC_WRITE_FAILED: '+str(exc)) from exc
        if failures:
            message = 'EXCEL_CLEANUP_FAILED: '+str(directory/'END.json')
            if original is not None:
                original.add_note(message)
            else:
                raise Stop(message)
