"""Separate candidate build; never replaces an existing executable."""
import os,sys,hashlib,json
from pathlib import Path
ROOT=Path(__file__).resolve().parents[2];sys.path.insert(0,str(ROOT))
from production_app import VERSION
assert VERSION=='1.0.0-rc1'
assert not (ROOT/'dist/production100').exists()
assert not (ROOT/'build/production100').exists()
paths=[ROOT/'production.py',*(ROOT/'production_app').glob('*.py'),ROOT/'stage9_app/supervisor.py']
for path in paths:compile(path.read_bytes(),str(path),'exec')
pins={str(path.relative_to(ROOT)):hashlib.sha256(path.read_bytes()).hexdigest() for path in paths}
sys.path.insert(0,str(ROOT/'tests/stage9/build_runtime'));import sitecustomize
os.environ['PYTHONPATH']=str(ROOT/'tests/stage9/build_runtime')
from PyInstaller.__main__ import run
os.chdir(ROOT)
run(['--distpath','dist/production100','--workpath','build/production100','production_100.spec'])
(ROOT/'validation/production_100/BUILD_INPUTS.json').write_text(json.dumps({'version':VERSION,'files':pins},indent=2),encoding='utf8')
