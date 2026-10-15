"""Rerun only two corrected harness checks; retain all initial evidence."""
from pathlib import Path
from copy import deepcopy
from xml.etree import ElementTree as ET
import ast
import hashlib
import json
import re
import shutil
import sys
import traceback
from datetime import datetime,timezone
OUT=Path(__file__).resolve().parent;ROOT=OUT.parents[1]
FIX=ROOT/'validation/stage4_excel_table_fix_20261007_121500'
sys.path.insert(0,str(ROOT));sys.path.insert(0,str(OUT));sys.dont_write_bytecode=True
sys.stdout.reconfigure(encoding='utf-8')
import zipfile
import profiles_stage4.storage as s
from native_probe import run as native
def read(p):return json.loads(Path(p).read_text(encoding='utf-8-sig'))
def sha(p):return hashlib.sha256(Path(p).read_bytes()).hexdigest()
def save(n,x):(OUT/n).write_text(json.dumps(x,ensure_ascii=False,indent=2),encoding='utf-8')
def require(x,m='assertion failed'):
    if not x:raise AssertionError(m)
CURRENT=s.DEFAULT_WORKBOOK;BASE=FIX/'baseline_generation/Профили.xlsx';CAT=s.load_catalog()
MAIN='http://schemas.openxmlformats.org/spreadsheetml/2006/main';NS={'m':MAIN}
source=OUT/'review.py';source_hash=sha(source)
tree=ast.parse(source.read_text(encoding='utf-8'))
selected=[n for n in tree.body if isinstance(n,ast.FunctionDef) and n.name in ['header_only','headers']]
assert len(selected)==2
exec(compile(ast.Module(body=selected,type_ignores=[]),str(source),'exec'),globals())
previous=read(OUT/'initial_harness_results.json')
corrections={'native root-cause A/B exact lexical header-only change':header_only,
             'actual renderer exact case empty invalid and valid header controls':headers}
assert {t['name'] for t in previous['tests'] if t['status']=='FAIL'}==set(corrections)
retests=[]
for name,fn in corrections.items():
    try:
        result=fn();retests.append({'name':name,'status':'PASS','detail':result});print('PASS '+name,flush=True)
    except BaseException:
        retests.append({'name':name,'status':'FAIL','traceback':traceback.format_exc()})
        save('corrected_harness_checks.json',retests);raise
save('corrected_harness_checks.json',retests)
assert sha(source)==source_hash
fix=read(FIX/'REVIEW_MANIFEST.json')
before=read(OUT/'preservation_before.json')
current={p:sha(p) for p in before}
assert before==read(OUT/'preservation_after.json')==current
save('preservation_final.json',current)
tests=[next((r for r in retests if r['name']==t['name']),t) for t in previous['tests']]
assert all(t['status']=='PASS' for t in tests) and len(tests)==11
result={'verdict':'PASS','passed':len(tests),'failed':0,'tests':tests,'DEV_artifacts':fix['DEV_artifacts'],
        'product_hashes':fix['product_hashes'],'changed_product_set':['profiles_stage4/storage.py','profiles_stage4/workbook.mjs'],
        'findings':{'BLOCKER':0,'MAJOR':0,'MINOR':0},'required_fix':'NONE','ready_to_resume_user_E2E':True,
        'initial_harness_corrections':[{'check':'root cause A/B','cause':'Validator regex did not handle x: namespace prefix','product_defect':False},
                                       {'check':'valid tiny export','cause':'Fixture truncated fixed-size UZK rules to two rows; kept all four for correct expectation','product_defect':False}],
        'not_run':previous['not_run']}
save('WORK_REREVIEW.json',result)
evidence={str(p.relative_to(OUT)):sha(p) for p in OUT.rglob('*') if p.is_file() and p.name!='REVIEW_MANIFEST.json'}
save('REVIEW_MANIFEST.json',{'created_at':datetime.now(timezone.utc).isoformat(),'verdict':'PASS','passed':11,'failed':0,
     'product_hashes':fix['product_hashes'],'DEV_artifacts':fix['DEV_artifacts'],'preservation_BEFORE_AFTER_CURRENT':True,
     'preserved_files':len(current),'evidence_files':evidence,'findings':result['findings'],'ready_to_resume_user_E2E':True})
print(json.dumps({'verdict':'PASS','passed':11,'failed':0,'preserved_files':len(current),'findings':result['findings']}),flush=True)
