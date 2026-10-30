import os,sys,ast
from pathlib import Path
ROOT=Path(__file__).resolve().parents[2]
for path in [ROOT/'stage9.py',*(ROOT/'stage9_app').glob('*.py')]:ast.parse(path.read_text(encoding='utf-8-sig'),filename=str(path))
sys.path.insert(0,str(ROOT/'tests/stage9/build_runtime'))
import sitecustomize,platform,faulthandler
def unavailable(*args,**kwargs):raise OSError('Build host WMI unavailable; native fallback')
platform._wmi_query=unavailable
faulthandler.dump_traceback_later(30,repeat=True)
os.environ['PYTHONPATH']=str(ROOT/'tests/stage9/build_runtime')
assert not (ROOT/'stage9_dist_092').exists(),'DO_NOT_OVERWRITE_EXISTING_BUILD'
from PyInstaller.__main__ import run
run(['--noconfirm','--distpath','stage9_dist_092','--workpath','stage9_build_092','stage9_092.spec'])
