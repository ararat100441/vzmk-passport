"""Excel ownership is local to Stage 8; keep COM apartment alive per thread.

Office proxy objects may outlive a Python context's local Excel variable.
Repeated CoUninitialize while such proxies exist can invalidate type-info
objects. Initialize once per calling thread; the OS releases it on process exit.
Every owned Excel instance is still explicitly closed and Quit in finally.
"""
from contextlib import contextmanager
import threading
from .storage import Stop

_apartment=threading.local()


@contextmanager
def excel_session():
    import pythoncom
    import win32com.client
    if not getattr(_apartment,'initialized',False):
        pythoncom.CoInitialize()
        _apartment.initialized=True
    excel=None
    try:
        excel=win32com.client.DispatchEx('Excel.Application')
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
        if excel is not None:
            try:
                excel.CutCopyMode=False
                excel.Workbooks.Close()
                excel.Quit()
            except Exception:
                pass
