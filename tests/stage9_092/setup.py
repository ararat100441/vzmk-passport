"""Build isolated inputs from already sealed DEV evidence; never read working registry."""
import sys,os,json,shutil,hashlib
from pathlib import Path
from uuid import uuid4
ROOT=Path(__file__).resolve().parents[2];sys.path.insert(0,str(ROOT))
from stage8_closure import storage as S
from stage9_app.workspace import save
from akz_stage6.core import digest as akz_digest
OUT=ROOT/'validation/stage9_092_20261012_01'
seed=Path(r'C:\Users\PUBLIC_USER\AppData\Local\Packages\OpenAI.Codex_2p2nqsd0c76g0\LocalCache\Local\VZMKWork09_fa550e56\VZMKPassport\workspaces\64cc6c3c7a0443b68da736b63a8329ee')
meta=S.read(seed/'STAGE9_WORKSPACE.json');assert meta['seal']==S.digest({k:v for k,v in meta.items() if k!='seal'})
base=Path(os.environ['LOCALAPPDATA'])/('VZMKStage9_092_'+uuid4().hex[:8]);base.mkdir()
source=base/'source';source.mkdir();copied={}
for rel,ref in meta['inputs'].items():
    src=seed/rel;assert S.sha(src)==ref['sha256'],str(src)
    dst=source/rel;dst.parent.mkdir(parents=True,exist_ok=True);shutil.copyfile(src,dst)
    assert S.sha(dst)==ref['sha256'];copied[rel]={'seed':str(src),'sha256':ref['sha256']}
save(source/'settings.json',{'documents_root':'documents','registry_path':'DEV/ВЗМК_DEV.xlsx','projects':{'Гараж Казахстан':{'passport_root':'project/Паспорт','template':'project/Паспорт/1_2026 - Приложение № 1.xlsx'}}})
# Only the assignment reference is remapped in an isolated copy of the accepted preview.
preview=S.read(ROOT/'validation/stage6_fix_r1_r2_20261009_094707/current_preview.json')
assert preview['seal']==akz_digest(preview['payload'])
keys=[p for p in preview['payload']['sources'] if 'Stage6_Fixtures' in p and p.endswith('.xlsx')];assert len(keys)==1
old=keys[0];assignment=source/'DEV/Stage6_Fixtures/assignments.xlsx';assert S.sha(assignment)==preview['payload']['sources'][old]
preview['payload']['sources'][str(assignment)]=preview['payload']['sources'].pop(old);preview['seal']=akz_digest(preview['payload'])
save(source/'validation/stage6_fix_r1_r2_20261009_094707/current_preview.json',preview)
old_e2e=S.read(ROOT/'validation/stage9_091_20261012_01/EXE_E2E_FINAL.json')
manifest=Path(old_e2e['manifests'][0]);vom=manifest.parent/'SOURCE_VOM.xls'
plan=S.read(manifest.parent/'PREVIEW.json');assert S.sha(vom)==plan['fingerprints'][str(vom)]
target=source/'TEST_VOM.xls';shutil.copyfile(vom,target);assert S.sha(target)==S.sha(vom)
# Make a probe-only clone. New import and all generation use the isolated source above.
work=base/'workspaces'/uuid4().hex;work.mkdir(parents=True)
for rel in meta['inputs']:
    dst=work/rel;dst.parent.mkdir(parents=True,exist_ok=True);shutil.copyfile(source/rel,dst)
(work/'validation').mkdir();newmeta=dict(meta,workspace_id=work.name);newmeta.pop('seal');newmeta['inputs']={r:{'source':str(source/r),'sha256':v['sha256']} for r,v in meta['inputs'].items()}
save(work/'STAGE9_WORKSPACE.json',newmeta|{'seal':S.digest(newmeta)})
import win32api,win32process,win32event
from stage9_app.ownership import process_image
users={}
for pid in win32process.EnumProcesses():
    try:
        h=win32api.OpenProcess(0x100400,False,pid)
        try:
            image=process_image(h)
            if Path(image).name.upper() in ('EXCEL.EXE','WINWORD.EXE') and win32event.WaitForSingleObject(h,0)==258:
                users[str(pid)]={'creation_time':win32process.GetProcessTimes(h)['CreationTime'].isoformat(),'image':image}
        finally:h.Close()
    except Exception:pass
save(OUT/'SETUP.json',{'base':str(base),'source':str(source),'probe_workspace':str(work),'vom':str(target),'copied':copied,'seed_manifest':str(seed/'STAGE9_WORKSPACE.json'),'user_office':users,'working_registry_used':False})
print(base,flush=True)
