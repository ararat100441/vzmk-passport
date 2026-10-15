import json,hashlib,sys,os
from pathlib import Path
from concurrent.futures import ThreadPoolExecutor
ROOT=Path(__file__).resolve().parents[2];OUT=Path(__file__).resolve().parent;OLD=ROOT/'validation/stage9_091_20261012_01'
def read(p):return json.loads(Path(p).read_text(encoding='utf-8-sig'))
def sha(p):return hashlib.sha256(Path(p).read_bytes()).hexdigest()
def save(n,d):
    with (OUT/n).open('x',encoding='utf8') as f:json.dump(d,f,ensure_ascii=False,indent=2)
m=read(OLD/'EVIDENCE_MANIFEST.json');refs={str(OLD/p):h for p,h in m['files'].items()}|m['external_files']
def verify(ref):
    p,h=ref;actual=sha(p) if Path(p).is_file() else None
    return None if actual==h else {'path':p,'expected':h,'actual':actual}
with ThreadPoolExecutor(max_workers=8) as pool:bad=[v for v in pool.map(verify,refs.items()) if v]
e2e=read(OLD/'EXE_E2E_FINAL.json');r=read(e2e['evidence']);steps=[]
for step in r['steps']:
    d=read(step['result']);s=read(d['office_supervisor'])
    steps.append({'task':step['task'],'expected':step['expected'],'actual':d['status'],'orphans':s['owned_orphans'],'ownership_errors':s['ownership_errors'],'result':step['result']})
manifests=[]
for p in r['manifests']:
    d=read(p);records=d['records'];unique={v['path'] for v in records}
    manifests.append({'path':p,'records':len(records),'unique_files':len(unique),'unique_word_files':len([v for v in unique if v.endswith('.docx')]),'unique_excel_files':len([v for v in unique if v.endswith('.xlsx')]),'passport':d['passport']})
save('HANDOFF_EVIDENCE_READBACK.json',{'status':'PASS' if not bad else 'FAIL','files_checked':len(refs),'mismatches':bad,'steps':steps,'manifests':manifests,'scope':'Readback of supplied historic evidence, not independent EXE E2E'})
print('HANDOFF',len(refs),'mismatches',len(bad),'steps',len(steps),flush=True)
