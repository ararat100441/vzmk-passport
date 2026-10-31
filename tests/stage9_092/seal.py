"""Seal exact, explicitly bounded TEST evidence, with no production traversal."""
from common import *
from stage8_closure.storage import digest
roots=[OUT,Path(SETUP['base']),WORKSPACES]
files={};counts={}
excluded={'EVIDENCE_MANIFEST.json','EVIDENCE_MANIFEST_SHA256.json','SEAL_VERIFY.json','SEAL.log'}
for root in roots:
    root=root.resolve();assert root!=ROOT and root!=ROOT.parent and root!=ROOT.parent.parent
    count=0
    for p in root.rglob('*'):
        if not p.is_file() or (p.parent==OUT and p.name in excluded):continue
        assert not p.is_symlink() and p.resolve().is_relative_to(root)
        files[str(p.resolve())]=sha(p);count+=1
    counts[str(root)]=count
manifest={'schema':'STAGE9_WORK09_TARGETED_HANDOFF_V1','version':'0.9.7','exe':str(EXE),'exe_sha256':sha(EXE),
    'source_inventory':str(OUT/'SOURCE_SHA256_FINAL.json'),'build_inventory':str(OUT/'BUILD_SHA256_FINAL.json'),
    'files':files,'aggregate_sha256':digest(files),'root_counts':counts,
    'working_registry_included':False,'independent_review':'NOT_RUN','ready_for_user_e2e':False,'ready_for_production':False}
save(OUT/'EVIDENCE_MANIFEST.json',manifest)
save(OUT/'EVIDENCE_MANIFEST_SHA256.json',{'path':str(OUT/'EVIDENCE_MANIFEST.json'),'sha256':sha(OUT/'EVIDENCE_MANIFEST.json')})
bad=[p for p,h in files.items() if sha(p)!=h];assert not bad,bad
save(OUT/'SEAL_VERIFY.json',{'status':'PASS','files':len(files),'sha_mismatch':len(bad),'manifest_sha256':sha(OUT/'EVIDENCE_MANIFEST.json'),'aggregate_sha256':digest(files)})
print('Sealed',len(files),'TEST evidence files; SHA mismatch 0; production excluded')
