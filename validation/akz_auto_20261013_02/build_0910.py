from common import *
import importlib.metadata,platform
runtime=[ROOT/'stage9.py',ROOT/'stage8.py',ROOT/'stage7_quality/render_emf.ps1']
for package in ('stage9_app','passport_app','profiles_stage4','uzk_stage5','akz_stage6','stage7_quality','stage8_closure'):runtime+=sorted((ROOT/package).glob('*.py'))
assert not EXE.parent.exists(),'NEW_BUILD_ONLY'
assert all(sha(p)==h for p,h in read(OUT/'ACCEPTED_BEFORE.json').items()),'ACCEPTED_SOURCE_CHANGED'
save(OUT/'SOURCE_AT_BUILD.json',{str(p):sha(p) for p in runtime+[ROOT/'stage9_0910.spec',ROOT/'tests/stage9_akz_auto/test_auto.py',ROOT/'tests/stage9_098/test_cleanup.py']})
dependencies={d.metadata['Name']:d.version for d in importlib.metadata.distributions() if d.metadata['Name']}
save(OUT/'BUILD_ENVIRONMENT.json',{'python':str(PY),'python_version':sys.version,'dependencies':dependencies,'build_hook':str(ROOT/'tests/stage9/build_runtime/sitecustomize.py'),'build_hook_sha256':sha(ROOT/'tests/stage9/build_runtime/sitecustomize.py'),'hook_in_product':False})
cmd=[str(PY),'-B','-m','PyInstaller','--noconfirm','--clean','--distpath',str(ROOT/'stage9_dist_0910'),'--workpath',str(ROOT/'stage9_build_0910'),str(ROOT/'stage9_0910.spec')]
env=os.environ.copy();env['PYTHONPATH']=str(ROOT/'tests/stage9/build_runtime');env['PYTHONUTF8']='1'
save(OUT/'BUILD_COMMAND.json',{'command':cmd,'PYTHONPATH':env['PYTHONPATH'],'started':time.time()})
with (OUT/'BUILD_0910.log').open('xb') as f:p=subprocess.run(cmd,env=env,stdout=f,stderr=subprocess.STDOUT,timeout=900)
assert p.returncode==0,(p.returncode,str(OUT/'BUILD_0910.log'))
assert all(sha(p)==h for p,h in read(OUT/'SOURCE_AT_BUILD.json').items()),'SOURCE_CHANGED_DURING_BUILD'
assert EXE.is_file();print('BUILD 0.9.10 PASS',sha(EXE),flush=True)
