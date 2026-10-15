"""Append-only Stage 9 diagnostics; accepted document receipts are unrelated."""
from pathlib import Path
import traceback
from uuid import uuid4
from .workspace import save

def error_record(step, exc):
    return {'step':step,'type':type(exc).__name__,'error':str(exc),
            'hresult':getattr(exc,'hresult',None),'excepinfo':getattr(exc,'excepinfo',None),
            'traceback':''.join(traceback.format_exception(exc))}

def viewer_diagnostic(folder):
    folder=Path(folder);folder.mkdir(parents=True,exist_ok=True)
    for _ in range(8):
        attempt=folder/('viewer_'+uuid4().hex)
        try:attempt.mkdir()
        except FileExistsError:continue
        return attempt/'ownership.json'
    raise RuntimeError('DIAGNOSTIC_ID_ALLOCATION_FAILED')

def event(diagnostic, kind, data):
    diagnostic=Path(diagnostic)
    for _ in range(8):
        path=diagnostic.with_name(diagnostic.stem+'_'+kind+'_'+uuid4().hex+'.json')
        try:save(path,data)
        except FileExistsError:
            if not path.is_file():raise
            continue  # Allocate another name; never overwrite existing evidence.
        return str(path)
    raise RuntimeError('DIAGNOSTIC_EVENT_ALLOCATION_FAILED: '+str(diagnostic))

class OfficeCleanupError(RuntimeError):
    def __init__(self, diagnostic, errors):
        self.diagnostic=str(diagnostic);self.cleanup_errors=list(errors)
        super().__init__('OFFICE_CLEANUP_FAILED: '+str(diagnostic)+'; '+
                         '; '.join(e['step']+': '+e['error'] for e in errors))
