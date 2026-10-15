import hashlib,json,sys
from pathlib import Path
RUN=Path(__file__).resolve().parent;ROOT=RUN.parents[1]
def read(p):return json.loads(Path(p).read_text(encoding='utf-8-sig'))
def sha(p):return hashlib.sha256(Path(p).read_bytes()).hexdigest()
def save(p,x):Path(p).write_text(json.dumps(x,ensure_ascii=False,indent=2),encoding='utf8')
cases=[read(p) for p in sorted((RUN/'cases').glob('*.json'))]
assert len(cases)==44
assert sum(x['status']=='PASS' for x in cases)==40
assert sum(x['status']=='FAIL' for x in cases)==4
integrity=read(RUN/'INTEGRITY_AFTER.json')
assert not integrity['review_changes'] and not integrity['protected_vs_checkpoint']
assert all(not g['mismatches'] for g in integrity['manifest_groups'].values())
summary={'status':'FAIL','total':44,'passed':40,'failed':4,'unique_findings':{'BLOCKER':0,'MAJOR':1,'MINOR':1},
 'gui_live':read(RUN/'GUI_LIVE_REVIEW.json')['status'],'protected_source_files':9301,'protected_changes':0,
 'stage8_product_changes':0,'official_closure':'NOT_RUN','real_uzk_decrements':0,'ready_for_user_e2e':'NO','cases':cases}
save(RUN/'REVIEW_SUMMARY.json',summary)
if not (RUN/'FINAL_REPORT.md').is_file():print('Summary saved; write FINAL_REPORT.md before sealing.');sys.exit()
source_manifest=read(ROOT/'validation/stage8_20261010_03/REVIEW_MANIFEST.json')
manifest={'schema':'STAGE8_INDEPENDENT_REVIEW_MANIFEST_V1','date':'2026-10-10','timezone':'Europe/Samara',
 'verdict':'FAIL','ready_for_user_e2e':'NO','product':source_manifest['product'],
 'prior_manifest_sha256':sha(ROOT/'validation/stage8_20261010_03/REVIEW_MANIFEST.json'),
 'review_evidence':{str(p.resolve()):sha(p) for p in sorted(RUN.rglob('*')) if p.is_file() and p.name not in ('REVIEW_MANIFEST.json','SEAL_VERIFICATION.json') and '__pycache__' not in p.parts}}
save(RUN/'REVIEW_MANIFEST.json',manifest)
checked=0;mismatches=[]
for group in ('product','review_evidence'):
    for path,h in manifest[group].items():
        checked+=1
        if sha(path)!=h:mismatches.append(path)
save(RUN/'SEAL_VERIFICATION.json',{'status':'PASS' if not mismatches else 'FAIL','checked':checked,'mismatches':mismatches,'manifest_sha256':sha(RUN/'REVIEW_MANIFEST.json')})
print(json.dumps({'summary':{k:v for k,v in summary.items() if k!='cases'},'sealed_files':len(manifest['review_evidence']),'checked':checked,'mismatches':mismatches},ensure_ascii=False))
