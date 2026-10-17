"""Build a separate 0.9.16 DEV/TEST distribution without overwriting old ones."""
import os
from pathlib import Path
import sys

ROOT=Path(__file__).resolve().parents[2]
sys.path.insert(0,str(ROOT))
from stage9_app import VERSION
assert VERSION=='0.9.16'
assert not (ROOT/'stage9_dist_0916').exists(),'DO_NOT_OVERWRITE_EXISTING_BUILD'
assert not (ROOT/'stage9_build_0916').exists(),'DO_NOT_OVERWRITE_EXISTING_BUILD'
for path in [ROOT/'stage9.py',*(ROOT/'stage9_app').glob('*.py')]:
    compile(path.read_bytes(),str(path),'exec')
sys.path.insert(0,str(ROOT/'tests/stage9/build_runtime'))
import sitecustomize
os.environ['PYTHONPATH']=str(ROOT/'tests/stage9/build_runtime')
from PyInstaller.__main__ import run
os.chdir(ROOT)
run(['--noconfirm','--distpath','stage9_dist_0916','--workpath','stage9_build_0916','stage9_0916.spec'])
