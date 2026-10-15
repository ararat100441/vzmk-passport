"""Continue existing 0.9.9; never launches or controls a GUI."""
from pathlib import Path
import json,os,hashlib,shutil
from uuid import uuid4
ROOT=Path(__file__).resolve().parent.parent
OLD=ROOT/'validation/stage9_fix_098_20261012_01'
OUT=ROOT/'validation/stage9_resume_099_20261012_01';OUT.mkdir()
def read(p):return json.loads(Path(p).read_text(encoding='utf-8-sig'))
def save(p,d):
    with Path(p).open('x',encoding='utf8') as f:json.dump(d,f,ensure_ascii=False,indent=2)
def sha(p):return hashlib.sha256(Path(p).read_bytes()).hexdigest()
prior=read(OLD/'SETUP_099.json');tag=uuid4().hex[:8]
base=Path(os.environ['LOCALAPPDATA'])/('VZMK099Resume_'+tag)
temp=base/'tmp';temp.mkdir(parents=True)
parent=Path.home()/('V9R099_'+tag);parent.mkdir()
tests=base/'source_tests';tests.mkdir();assert tests.resolve().is_relative_to(base.resolve())
setup={'base':str(base),'temp':str(temp),'parent':str(parent),'source':prior['source'],'vom':prior['vom'],
       'source_test_parent':str(tests),'gui_permission':False,'prior_evidence':str(OLD),
       'preserve_preexisting_gui':read(OLD/'TARGETED_GUI_LAUNCH.json')}
save(OUT/'SETUP.json',setup)
provenance=[]
for name in ('common.py','lifecycle.py','e2e.py','negative.py','normal_overlap.py','overlap_crash_cleanup.py',
             'ownership_stop.py','foreign_pumped.py','window_close_guard.py','extended_guard.py','semantic_audit.py','receipts.py'):
    source=OLD/name;before=source.read_text(encoding='utf8');after=before
    if name=='common.py':
        after=after.replace("OUT/'SETUP_099.json'","OUT/'SETUP.json'").replace("OUT/'NEW_EXE_BINDING.json'","Path(SETUP['prior_evidence'])/'NEW_EXE_BINDING.json'")
        after=after.replace('Independent frozen-task controller','Codex self-test frozen-task controller')
    if name=='e2e.py':
        line="    smoke=run('smoke',workspace=work);assert smoke['frozen'] and smoke['python_in_path'] is None and smoke['version']=='0.9.9'"
        assert line in after
        after=after.replace(line,"    # GUI-producing smoke omitted by explicit user STOP; no Tk is created.")
        after=after.replace("    run('smoke',workspace=work)","    # GUI relaunch smoke is NOT_RUN pending separate permission.")
        after=after.replace("'python_in_path':None}","'python_in_path':None,'gui_smoke':'NOT_RUN_USER_STOP','scope':'FROZEN_CLI_ONLY'}")
    after=after.replace('Independent output audit','Codex output audit').replace('Independent manifest seals','Codex manifest seals')
    after=after.replace('INDEPENDENT_DOCUMENT_AND_XLSM_AUDIT.json','DOCUMENT_AND_XLSM_AUDIT.json')
    (OUT/name).write_text(after,encoding='utf8')
    provenance.append({'name':name,'from':str(source),'before_sha256':sha(source),'adapted_sha256':sha(OUT/name),'adaptation':'fresh own roots/results; no GUI smoke; self-test labels'})
save(OUT/'CONTROLLER_PROVENANCE.json',provenance)
print(OUT)
