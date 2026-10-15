"""Actual frozen CLI repeated read-only opens; does not call GUI Viewer."""
from common import *
e2e=read(OUT/'EXE_E2E_FINAL.json');assert e2e['exe_sha256']==sha(EXE)
RUN=OUT/('repeat_native_'+uuid4().hex[:8]);RUN.mkdir();cases=[];paths=e2e['manifests']
before={r['path']:sha(r['path']) for m in paths for r in read(m)['records']}
for i in (0,0,1,0):
    d,sup,p=task(RUN,'native-verify',workspace=e2e['workspace'],manifest=paths[i])
    native=read(d['evidence']);assert native['physical_files']==6 and native['document_records']==7
    assert all(r['readonly'] is True for r in native['word']+native['excel'])
    assert all(sha(p)==h for p,h in before.items())
    cases.append({'manifest_index':i,'result':str(p),'native_evidence':d['evidence'],'supervisor':d['office_supervisor']})
assert len({c['native_evidence'] for c in cases})==4
save(RUN/'RESULT.json',{'status':'PASS','exe_sha256':sha(EXE),'cases':cases,'all_document_SHA_unchanged':True,'gui_viewer':'NOT_RUN_USER_STOP'})
save(OUT/'REPEATED_NATIVE_OPEN_FINAL.json',{'status':'PASS','exe_sha256':sha(EXE),'evidence':str(RUN/'RESULT.json'),'cycles':4,'GUI':False})
