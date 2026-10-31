"""Seal the original successful gates plus the corrected targeted fault test."""
import sys
from pathlib import Path
ROOT=Path(__file__).resolve().parents[2];sys.path.insert(0,str(ROOT))
from production_app.storage import read,save,sha
from production_app.core import pending

def main():
    out=Path(r'C:\VZMKPassportChecks\20261018\frozen-final')
    initial=read(out/'progress.json')
    assert all(r['exit_code']==0 for r in initial if r['label']!='frozen_fault_recovery')
    fault=read(out/'fault-corrected/RESULT.json');e2e=read(out/'e2e/RESULT.json')
    gui=read(out/'gui-components/RESULT.json');visual=read(out.parent/'gui-native/VISUAL_RESULT.json')
    for r in (fault,e2e,gui,visual,read(out/'observer_cleanup_result.json')):assert r['status']=='PASS'
    assert not pending(e2e['cfg']['registry'])
    inputs=read(ROOT/'validation/production100-final/BUILD_INPUTS.json')
    assert all(sha(ROOT/path)==h for path,h in inputs['files'].items()),'BUILD INPUT DRIFT'
    exe=Path(r'C:\VZMK-Passport\1.0.0-rc3\VZMK-Passport-1.0.0-rc3.exe')
    assert sha(exe)==fault['exe_sha256']
    save(out/'RESULT.json',{'status':'PASS','exe':str(exe),'exe_sha256':sha(exe),
        'initial_gates':initial,'fault_original_attempt':'INVALID_TEST_OBSERVER; original registry bytes retained; copied operation recovered',
        'targeted_fault_result':str(out/'fault-corrected/RESULT.json'),'targeted_fault':'PASS_AFTER_OBSERVER_REPAIR',
        'normal_frozen_e2e':'PASS','gui_components':'PASS','native_gui_startup':'PASS; computer use later stopped by user Escape',
        'frozen_without_python_path':True,'actual_issue':'NOT_RUN','production_registry_written':False,
        'build_inputs_unchanged':True,'release_readiness':'BLOCKED_REAL_TEMPLATE_COMPATIBILITY_AND_SHIPMENT_INPUTS'})
    print('TECHNICAL_GATES_PASS; OPERATIONAL_1.0.0_BLOCKED')
if __name__=='__main__':main()
