"""Targeted repeated Office ownership check after the COM lifecycle correction."""
import json
from pathlib import Path
import sys
from stage8_closure import core,storage as S
from stage8_closure.fixtures import create

run=Path(sys.argv[1]).resolve()
before=core.code_bindings();rounds=[]
for i in range(5):
    f=create(run/'sandboxes');p=core.preview(f['sandbox'],f['request'])
    r=core.close(p,core.confirmation(p))
    assert r['status']=='CLOSED_TEST'
    assert core.recover(f['sandbox'])['status']=='NO_PENDING_OPERATION'
    rounds.append({'round':i+1,'sandbox':f['sandbox'],'status':r['status']})
    print('Office stability round',i+1,'PASS',flush=True)
assert before==core.code_bindings()
S.write(run/'OFFICE_STABILITY.json',{'status':'PASS','rounds':rounds,'product':before,
    'purpose':'Repeated native creation, save, validation and close in one process; fresh Excel instances.'})
