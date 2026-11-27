"""Optional progress messages within a private GUI job."""
from pathlib import Path

from .storage import local_root, save

_output=None


def bind(request):
    global _output
    _output=None
    if request:
        path=Path(request)
        if path.name=='request.json' and path.parent.parent==local_root()/'gui-jobs':
            _output=path.parent/'progress.json'


def report(message):
    if _output is not None:
        try:save(_output,{'message':message})
        except OSError:pass  # Progress must never interrupt a transaction.
