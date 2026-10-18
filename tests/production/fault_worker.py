"""TEST ONLY: simulate loss of control immediately after Windows ReplaceFile.

Uses the actual supervisor, original Office guard and actual NTFS replacement.
No fault switch is shipped in the production GUI or executable.
"""
import sys
from pathlib import Path
ROOT=Path(__file__).resolve().parents[2];sys.path.insert(0,str(ROOT))
import production
from production_app import core

if '--office-cleanup' in sys.argv:
    from stage9_app.supervisor import cleanup_entry
    sys.exit(cleanup_entry(sys.argv[sys.argv.index('--office-cleanup')+1]))
if '--office-worker' not in sys.argv:
    from stage9_app.supervisor import supervise
    sys.exit(supervise(sys.argv[1:],entrypoint=Path(__file__)))
original=core.atomic_replace
def crash(source,target,**kwargs):
    original(source,target,**kwargs)
    raise RuntimeError('TEST_CRASH_AFTER_REPLACEFILE_BEFORE_JOURNAL_UPDATE')
core.atomic_replace=crash
sys.exit(production.main())
