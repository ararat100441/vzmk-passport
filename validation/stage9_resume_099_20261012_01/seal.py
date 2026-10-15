"""Seal only this resume's immutable TEST evidence and explicit own roots."""
from common import *
from stage8_closure.storage import digest
roots=[OUT,Path(SETUP['base']),WORKSPACES];files={};counts={}
exclude={'EVIDENCE_MANIFEST.json','EVIDENCE_MANIFEST_SHA256.json','SEAL_VERIFY.json','SEAL.log'}
for root in roots:
    root=root.resolve();assert root not in (ROOT,ROOT.parent,ROOT.parent.parent)
    count=0
    for p in root.rglob('*'):
        if not p.is_file() or (p.parent==OUT and p.name in exclude):continue
        assert not p.is_symlink() and p.resolve().is_relative_to(root)
        files[str(p.resolve())]=sha(p);count+=1
    counts[str(root)]=count
prep=ROOT/'validation/resume099_prepare.py';files[str(prep)]=sha(prep)
state=read(OUT/'REVIEW_STATUS.json')
save(OUT/'EVIDENCE_MANIFEST.json',{'schema':'STAGE9_099_RESUME_SELF_TEST_V1','version':'0.9.9','exe':str(EXE),'exe_sha256':sha(EXE),
    'files':files,'aggregate_sha256':digest(files),'root_counts':counts,'build_inventory':str(OUT/'BUILD_SHA256_AFTER.json'),
    'source_inventory':str(OUT/'SOURCE_SHA256_AFTER.json'),'review_status':state,'working_registry_included':False,
    'preexisting_GUI_root_excluded':True,'independent_review':'NOT_RUN','GUI':'NOT_RUN_USER_STOP'})
save(OUT/'EVIDENCE_MANIFEST_SHA256.json',{'path':str(OUT/'EVIDENCE_MANIFEST.json'),'sha256':sha(OUT/'EVIDENCE_MANIFEST.json')})
bad=[p for p,h in files.items() if sha(p)!=h];assert not bad,bad
save(OUT/'SEAL_VERIFY.json',{'status':'PASS','file_count':len(files),'sha_mismatch':0,'aggregate_sha256':digest(files),'manifest_sha256':sha(OUT/'EVIDENCE_MANIFEST.json')})
print('Sealed',len(files),'TEST files; mismatch 0; GUI NOT_RUN; production excluded')
