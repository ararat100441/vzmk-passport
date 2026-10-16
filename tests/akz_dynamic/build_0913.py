"""Build a new DEV/TEST distribution; refuse to overwrite an existing one."""
import ast
import os
from pathlib import Path
import sys

ROOT=Path(__file__).resolve().parents[2]
sys.path.insert(0,str(ROOT))
from stage9_app import VERSION
assert VERSION=='0.9.13'
assert not (ROOT/'stage9_dist_0913').exists(),'DO_NOT_OVERWRITE_EXISTING_BUILD'
assert not (ROOT/'stage9_build_0913').exists(),'DO_NOT_OVERWRITE_EXISTING_BUILD'
for folder in ('akz_stage6','stage9_app'):
    for path in (ROOT/folder).glob('*.py'):
        ast.parse(path.read_text(encoding='utf-8-sig'),filename=str(path))
sys.path.insert(0,str(ROOT/'tests/stage9/build_runtime'))
import sitecustomize
os.environ['PYTHONPATH']=str(ROOT/'tests/stage9/build_runtime')
from PyInstaller.__main__ import run
os.chdir(ROOT)
run(['--noconfirm','--distpath','stage9_dist_0913','--workpath','stage9_build_0913','stage9_0913.spec'])
