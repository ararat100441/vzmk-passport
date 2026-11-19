"""Manifest-last sealing of the finished handoff, without changing evidence."""
from common import *
from stage8_closure.storage import digest
assert read(OUT/'FINAL_AUDIT.json')['status']=='PASS'
assert (OUT/'STAGE9_REPORT.md').is_file() and (OUT/'WORK09_ASSIGNMENT.md').is_file()
target=OUT/'EVIDENCE_MANIFEST.json';assert not target.exists()
files={str(p.relative_to(OUT)):sha(p) for p in sorted(OUT.rglob('*')) if p.is_file()}
external={};jobs=set();workspaces=set()
def visit(value):
    if isinstance(value,dict):
        for v in value.values():visit(v)
    elif isinstance(value,list):
        for v in value:visit(v)
    elif isinstance(value,str) and len(value)<32760:
        p=Path(value)
        if not p.is_absolute():return
        try:
            if p.is_file() and not p.is_relative_to(OUT):external[str(p)]=sha(p)
            if p.name=='SUPERVISOR.json' and p.is_file():jobs.add(p.parent)
        except OSError:pass
for p in OUT.rglob('*.json'):
    try:visit(read(p))
    except (ValueError,OSError):pass
for name in ('LIFECYCLE_FINAL','EXE_E2E_FINAL'):
    workspaces.add(Path(read(OUT/(name+'.json'))['workspace']))
for folder in sorted(jobs|workspaces):
    for p in folder.rglob('*'):
        if p.is_file():external[str(p)]=sha(p)
external[str(ROOT/'STAGE9_DEV_091.md')]=sha(ROOT/'STAGE9_DEV_091.md')
save(target,{'schema':'STAGE9_091_EVIDENCE_V1','exe_sha256':sha(EXE),'source_inventory':str(OUT/'SOURCE_SHA256_FINAL.json'),'source_aggregate_sha256':read(OUT/'SOURCE_SHA256_FINAL.json')['aggregate_sha256'],'files':files,'external_files':external,'aggregate_sha256':digest({'files':files,'external_files':external}),'manifest_written_last':True,'historical_failures_retained':True,'independent_work_review':'NOT_RUN'})
print('SEALED',len(files),'local evidence files,',len(external),'referenced external files',sha(target))
