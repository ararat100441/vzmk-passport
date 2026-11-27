"""Separate candidate build; never replaces an existing executable."""
import os,sys,hashlib,json,argparse
from importlib.metadata import PackageNotFoundError, version
from pathlib import Path
ROOT=Path(__file__).resolve().parents[2];sys.path.insert(0,str(ROOT))
from production_app import VERSION
assert VERSION in ('1.0.0-rc3','1.0.0')
p=argparse.ArgumentParser();p.add_argument('--name',default='production100');args=p.parse_args()
assert args.name.replace('-','').isalnum()
assert not (ROOT/'dist'/args.name).exists()
assert not (ROOT/'build'/args.name).exists()
packages=['production_app','stage9_app','passport_app','profiles_stage4','uzk_stage5','akz_stage6','stage7_quality','stage8_closure']
locks=[ROOT/'requirements-win-py312.lock',ROOT/'requirements-build-win-py312.lock']
runtime_path=ROOT/'tests/stage9/build_runtime'
resolved={}
for lock in locks:
    for line in lock.read_text(encoding='utf8').splitlines():
        if not line.strip() or line.lstrip().startswith('#'):continue
        name,expected=line.split()[0].split('==',1)
        try:installed=version(name)
        except PackageNotFoundError:raise SystemExit(f'Install {name}=={expected} from {lock.name} before building.')
        if installed!=expected:
            raise SystemExit(f'{lock.name} requires {name}=={expected}; installed {installed}.')
        resolved[name]=installed
paths=[ROOT/'production.py',ROOT/'production_100.spec',ROOT/'docs/production_user_guide.txt',
       Path(__file__).resolve(),runtime_path/'sitecustomize.py',ROOT/'requirements.txt',*locks,
       *[path for package in packages for path in (ROOT/package).glob('*.py')]]
paths.sort(key=lambda path:path.relative_to(ROOT).as_posix())
for path in paths:
    if path.suffix=='.py':compile(path.read_bytes(),str(path),'exec')
pins={str(path.relative_to(ROOT)):hashlib.sha256(path.read_bytes()).hexdigest() for path in paths}
sys.path.insert(0,str(runtime_path));import sitecustomize
os.environ['PYTHONPATH']=str(runtime_path)
from PyInstaller.__main__ import run
os.chdir(ROOT)
run(['--distpath','dist/'+args.name,'--workpath','build/'+args.name,'production_100.spec'])
assert all(hashlib.sha256((ROOT/path).read_bytes()).hexdigest()==h for path,h in pins.items()),'BUILD INPUT DRIFT'
evidence=ROOT/'validation'/args.name/'BUILD_INPUTS.json';evidence.parent.mkdir(parents=True,exist_ok=True)
evidence.write_text(json.dumps({'version':VERSION,'python':sys.version,'packages':resolved,
                                'files':pins},indent=2),encoding='utf8')
