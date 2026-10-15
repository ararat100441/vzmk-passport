"""Crash only an artificial fixture using exact raw engine files shipped in EXE.

The imported __file__ paths match the frozen process's Stage 8 code bindings.
No pins, checks, templates, code files or business functions are replaced.
"""
import sys,os
from pathlib import Path
internal,workspace,preview=sys.argv[1:4]
sys.path.insert(0,internal)
from stage9_app.workspace import bind
from stage9_app.ownership import owned_office
from stage8_closure import core,storage as S
from profiles_stage4.offline import no_network
root,_=bind(workspace)
assert str(Path(core.__file__).resolve()).startswith(str(Path(internal).resolve()))
plan=S.read(preview)
def fault(point,path):
    if point=='during_publish':os._exit(73)
with no_network(),owned_office(Path(preview).with_name('CRASH_OFFICE.json')):
    core.close(plan,core.confirmation(plan),_fault=fault)
raise AssertionError('Expected process death')
