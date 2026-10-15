"""Targeted correction of independent probe's protected-file read, not product code."""
from pathlib import Path
from unittest.mock import patch
import hashlib,json,os,shutil,subprocess,sys
OUT=Path(__file__).resolve().parent;ROOT=OUT.parents[1]
sys.path.insert(0,str(ROOT));sys.dont_write_bytecode=True
sys.stdout.reconfigure(encoding='utf-8')
def read(p):return json.loads(Path(p).read_text(encoding='utf-8-sig'))
def sha(p):return hashlib.sha256(Path(p).read_bytes()).hexdigest()
def check(v,msg='Independent commit-probe assertion failed'):
    if not v:raise AssertionError(msg)
def audit(event,args):
    if event in ('socket.connect','socket.getaddrinfo','socket.sendto'):raise AssertionError('No network')
    if event in ('open','os.rename','os.remove','os.mkdir','os.rmdir'):
        paths=[a for a in args if isinstance(a,(str,bytes))]
        for a in paths:
            text=os.fsdecode(a).replace('\\','/')
            check('/Работа/ВЗМК.xlsx' not in text and '/Работа/Профили.xlsx' not in text,'No production')
        write=event!='open'
        if event=='open':
            mode,flags=args[1:3]
            write=bool(mode and any(c in mode for c in 'wax+')) or bool(flags & (os.O_WRONLY|os.O_RDWR|os.O_CREAT|os.O_TRUNC|os.O_APPEND))
        if write:
            for a in paths[:2] if event=='os.rename' else paths[:1]:check(Path(os.fsdecode(a)).resolve().is_relative_to(OUT),'Write outside fixture')
sys.addaudithook(audit)
import profiles_stage4.storage as s
import profiles_stage4.windows_generation as wg
real=s.DEFAULT_WORKBOOK
names=[real.name,s.STATE,s.CATALOG,s.COMMIT]
product=read(ROOT/'validation/stage4_fix_m1_m2_20261007_090000/product_hashes_current.json')
check({p:sha(ROOT/p) for p in product}==product)
folder=OUT/'fixtures/commit_boundary_recheck_utf8/DEV/Справочники';folder.mkdir(parents=True)
for n in names:shutil.copyfile(real.parent/n,folder/n)
t=folder/real.name
previous={n:sha(folder/n) for n in names}
phases=[]
writer='''from pathlib import Path
import json,os,sys
sys.stdout.reconfigure(encoding='utf-8')
folder=Path(sys.argv[1]).resolve()
base=Path(sys.argv[2]).resolve()
assert folder.is_relative_to(base/'fixtures')
results=[]
for name in ['Профили.xlsx','profiles.user.json','profiles.catalog.json','profiles.build.json']:
 for op in ['write','delete','rename','replace']:
  p=folder/name;q=folder/(name+'.independent-attempt')
  if op=='replace':q.write_bytes(b'EXTERNAL_REPLACEMENT')
  try:
   if op=='write':p.write_bytes(b'EXTERNAL_WRITE')
   elif op=='delete':p.unlink()
   elif op=='rename':p.rename(q)
   else:os.replace(q,p)
   results.append({'file':name,'operation':op,'blocked':False})
  except OSError as e:results.append({'file':name,'operation':op,'blocked':True,'winerror':e.winerror,'errno':e.errno})
  finally:q.unlink(missing_ok=True)
print(json.dumps(results,ensure_ascii=False))
'''
(OUT/'boundary_writer.py').write_text(writer,encoding='utf-8')
def attacks(phase):
    completed=subprocess.run([sys.executable,'-I','-B',str(OUT/'boundary_writer.py'),str(folder),str(OUT)],capture_output=True,text=True,encoding='utf-8',check=True,timeout=25)
    result=json.loads(completed.stdout)
    check(len(result)==16)
    check(all(r['blocked'] and (r['winerror'] in (5,32,33) or (r['operation']=='write' and r['winerror'] is None and r['errno']==13)) for r in result),str(result))
    phases.append({'phase':phase,'attempts':result})
def renderer(payload,destination,previews=None):
    check(read(payload)==s.workbook_payload(s.load_catalog(real)),'Copied workbook payload must match the exact current workbook')
    shutil.copyfile(real,destination)
    return {'network_attempts':0}
publish=wg.GenerationGuard.publish;rename=wg.ProtectedFile.rename
def protected_publish(self,stage,fail_at=None):
    attacks('immediately_before_first_rename_all_previous_files_protected')
    return publish(self,stage,fail_at)
def protected_rename(self,destination):
    rename(self,destination)
    if Path(destination)==t: attacks('immediately_after_first_new_file_install_mid_commit')
    if Path(destination)==folder/s.COMMIT: attacks('immediately_after_manifest_install_before_handles_release')
with patch.object(s,'render_workbook',renderer),patch.object(wg.GenerationGuard,'publish',protected_publish),patch.object(wg.ProtectedFile,'rename',protected_rename):
    built=s.build(t)
check(len(phases)==3)
backup=Path(built['backup'])
check({n:sha(backup/n) for n in names}==previous)
s.validated_generation({n:(backup/n).read_bytes() for n in names},t.name)
s.load_catalog(t)
check({p:sha(ROOT/p) for p in product}==product)
result={'status':'PASS','separate_process':True,'phases':phases,'blocked_attempts':48,'protected_files':names,'backup_exact_previous_generation':True,'normal_commit':'PASS','product_before_after_current_equal':True,'real_DEV_builds':0,'supersedes':'Initial independent probe opened protected files with CRT share mode incompatible with DELETE-capable protection. This retry checks outcomes after release and observes independent writer OS denials.'}
(OUT/'commit_boundary_recheck.json').write_text(json.dumps(result,ensure_ascii=False,indent=2),encoding='utf-8')
print(json.dumps({'status':'PASS','blocked_attempts':48,'boundary_phases':3,'backup_coherent':True}),flush=True)
