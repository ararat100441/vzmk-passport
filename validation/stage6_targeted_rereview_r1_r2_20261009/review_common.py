from pathlib import Path
import sys, os, json, hashlib, traceback, shutil, copy, ast, difflib
from unittest.mock import patch
from decimal import Decimal
sys.stdout.reconfigure(encoding='utf-8')
E=Path(__file__).resolve().parent
ROOT=E.parents[1]
FIX=ROOT/'validation/stage6_fix_r1_r2_20261009_094707'
sys.path.insert(0,str(ROOT))
AUDIT={'production_attempts':[],'network_attempts':[]}
def auditor(event,args):
    if event=='open' and isinstance(args[0],(str,bytes,os.PathLike)):
        p=Path(os.fsdecode(args[0]))
        if p.name.casefold() in ('взмк.xlsx','узк.xlsm'):
            AUDIT['production_attempts'].append(str(p));raise AssertionError('PRODUCTION DENIED')
    if event in ('socket.connect','socket.getaddrinfo'):
        AUDIT['network_attempts'].append(str(args));raise AssertionError('NETWORK DENIED')
sys.addaudithook(auditor)
def load(p):return json.loads(Path(p).read_text('utf-8-sig'))
def save(n,v):(E/n).write_text(json.dumps(v,ensure_ascii=False,indent=2,default=str),encoding='utf-8')
def sha(p):return hashlib.sha256(Path(p).read_bytes()).hexdigest()
ACTNAME='ТЕСТ АКЗ №0.xlsx'
APPNAME='ТЕСТ Приложение к АКЗ №0.xlsx'
PAIR=FIX/'NEW_TEST_9_2026'
EXPECTED={ACTNAME:'b640796570a8d04b6bca404205d99eae4c7005c700587812ed25bc99e715eb1a',APPNAME:'7766c92df82df2484eb82cd6917769f5d9ede912b3f0e8f26e3f8aaabdd357f2','PAIR.json':'eaad2385ef401a21f9203db9648530f71939ec4933296f309adb7db19bedbb8f'}
