import os,sys
import ast
from pathlib import Path
HERE=Path(__file__).resolve().parent
for path in [HERE.parents[1]/'stage9.py',*(HERE.parents[1]/'stage9_app').glob('*.py')]:ast.parse(path.read_text(encoding='utf-8-sig'),filename=str(path))
sys.path.insert(0,str(HERE/'build_runtime'))
import sitecustomize
import platform,faulthandler
def unavailable(*args,**kwargs):raise OSError('Build host WMI unavailable; native fallback')
platform._wmi_query=unavailable
faulthandler.dump_traceback_later(30,repeat=True)
os.environ['PYTHONPATH']=str(HERE/'build_runtime')
from PyInstaller.__main__ import run
run(['--noconfirm','--distpath','stage9_dist_091','--workpath','stage9_build_091',sys.argv[1] if len(sys.argv)>1 else 'stage9_091.spec'])
