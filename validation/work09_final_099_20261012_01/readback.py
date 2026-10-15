from common import *
import concurrent.futures
rows=[]
for relative in ('validation/stage9_resume_099_20261012_01','validation/stage9_gui_099_20261012_01'):
    folder=ROOT/relative;p=folder/'EVIDENCE_MANIFEST.json';m=read(p);files=m['files']
    with concurrent.futures.ThreadPoolExecutor(10) as pool:
        bad=[r for r in pool.map(lambda item:None if Path(item[0]).is_file() and sha(item[0])==item[1] else item[0],files.items()) if r]
    rows.append({'manifest':str(p),'sha256':sha(p),'references':len(files),'mismatch':bad})
    assert not bad,bad
save(OUT/'HANDOFF_READBACK.json',{'status':'PASS','sets':rows,'runtime_proof':False})
print('Historical self-evidence readback only',[(x['references'],len(x['mismatch'])) for x in rows],flush=True)
