"""TEST ONLY: real sealed-plan cleanup, injected proof failure remains FAIL."""
from pathlib import Path
import sys
ROOT=Path(__file__).resolve().parents[2];sys.path.insert(0,str(ROOT))
from stage9_app import ownership
from stage9_app.supervisor import cleanup_entry
original=ownership.quit_owned_window
def failed_proof(*args,**kwargs):
    report=original(*args,**kwargs)
    report['status']='FAIL';report['error']='TEST_INJECTED_CLEANUP_PROOF_FAILURE'
    return report
ownership.quit_owned_window=failed_proof
sys.exit(cleanup_entry(sys.argv[sys.argv.index('--office-cleanup')+1]))
