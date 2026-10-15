import ast,hashlib,json,re
from datetime import datetime,timezone
from pathlib import Path
RUN=Path(__file__).resolve().parent;ROOT=RUN.parents[1]
FIX=ROOT/'validation/stage5_S5_R1_fix_20261008_115014'
PREV=ROOT/'validation/stage5_independent_review_20261008_01'
def read(p):return json.loads(Path(p).read_text(encoding='utf-8-sig'))
def sha(p):
    p=Path(p);assert p.name.casefold() not in ('взмк.xlsx','узк.xlsm'),'Forbidden production hash'
    return hashlib.sha256(p.read_bytes()).hexdigest()
def write(n,v):(RUN/n).write_text(json.dumps(v,ensure_ascii=False,indent=2),encoding='utf8')
m=read(FIX/'REVIEW_MANIFEST.json');old=read(PREV/'REVIEW_MANIFEST.json')
current={p:sha(p) for p in m['product_current']}
assert current==m['product_before_final_tests']==m['product_after']==m['product_current'],'STOP: current tested product mismatch'
bad=[]
for name,meta in m['evidence_files'].items():
    p=(FIX/name).resolve();assert p.is_relative_to(FIX)
    if sha(p)!=meta['sha256'] or p.stat().st_size!=meta['bytes']:bad.append(name)
assert not bad,('Fix evidence mismatch',bad)
actual={str(p.relative_to(FIX)) for p in FIX.rglob('*') if p.is_file() and p.name!='REVIEW_MANIFEST.json'}
assert actual==set(m['evidence_files'])
assert all((FIX/n).stat().st_mtime<=(FIX/'REVIEW_MANIFEST.json').stat().st_mtime for n in m['evidence_files'])
assert sha(PREV/'REVIEW_MANIFEST.json')==m['reviewed_baseline']['review_manifest_sha256']
assert all(sha(PREV/n)==meta['sha256'] and (PREV/n).stat().st_size==meta['bytes'] for n,meta in old['evidence_files'].items())
baseline=read(FIX/'baseline.json')['protected_before']
allowed={str(ROOT/'uzk_stage5'/n) for n in ('core.py','inputs.py','gui.py')}
changed={p for p,h in baseline.items() if sha(p)!=h}
assert changed==allowed,('Unexpected baseline changes',sorted(changed))
doc=m['existing_test_docx'];assert sha(doc['path'])==doc['sha256']=='7723be279f6e776a472fb33f0b1bc153aecd14fc04e5b319fae73f89f2960661'
assert sha(ROOT/'uzk_stage5/word.py')==old['review_source_before'][str(ROOT/'uzk_stage5/word.py')]

# Independently reverse the supplied patch and bind each recovered old source to
# the previous review SHA. This checks the complete byte scope, not a scope label.
lines=(FIX/'S5-R1.patch').read_text(encoding='utf-8-sig').splitlines()
recovered={};scopes={};i=0
while i<len(lines):
    if not lines[i].startswith('--- BEFORE/'):i+=1;continue
    rel=lines[i][11:];assert lines[i+1]=='+++ AFTER/'+rel;i+=2
    now=(ROOT/rel).read_text(encoding='utf8').splitlines();cursor=0;before=[]
    while i<len(lines) and not lines[i].startswith('--- BEFORE/'):
        if not lines[i].startswith('@@'):i+=1;continue
        h=re.match(r'@@ -(\d+)(?:,(\d+))? \+(\d+)(?:,(\d+))? @@',lines[i]);assert h
        start=int(h.group(3))-1;before.extend(now[cursor:start]);cursor=start;i+=1
        while i<len(lines) and not lines[i].startswith(('@@','--- BEFORE/')):
            s=lines[i];i+=1
            if not s or s.startswith('\\'):continue
            if s[0] in (' ','+'):
                assert now[cursor]==s[1:],(rel,cursor,now[cursor],s);cursor+=1
            if s[0] in (' ','-'):before.append(s[1:])
    before.extend(now[cursor:])
    expected=old['review_source_before'][str(ROOT/rel)]
    candidates=[sep.join(before)+end for sep in ('\n','\r\n') for end in ('','\n','\r\n')]
    match=next((v for v in candidates if hashlib.sha256(v.encode('utf8')).hexdigest()==expected),None)
    assert match is not None,('Patch does not explain all changes',rel)
    dest=RUN/'recovered_previous_source'/rel;dest.parent.mkdir(parents=True,exist_ok=True);dest.write_bytes(match.encode('utf8'))
    def functions(src):
        tree=ast.parse(src);out={}
        for n in tree.body:
            if isinstance(n,(ast.FunctionDef,ast.AsyncFunctionDef)):out[n.name]=ast.dump(n,include_attributes=False)
            if isinstance(n,ast.ClassDef):
                for f in n.body:
                    if isinstance(f,(ast.FunctionDef,ast.AsyncFunctionDef)):out[n.name+'.'+f.name]=ast.dump(f,include_attributes=False)
        return out
    a=functions(match);b=functions((ROOT/rel).read_text(encoding='utf8'))
    scopes[rel]=[k for k in a.keys()|b.keys() if a.get(k)!=b.get(k)]
    recovered[rel]=dict(previous_sha=expected,current_sha=sha(ROOT/rel),patch_fully_accounts_for_change=True)
assert scopes=={'uzk_stage5/core.py':['preview'],'uzk_stage5/inputs.py':['mark_profiles'],'uzk_stage5/gui.py':['App.poll']},scopes
assert {p.name for p in (ROOT/'uzk_stage5').glob('*.py')}=={'__init__.py','core.py','inputs.py','gui.py','word.py'}
observed={**{p:sha(p) for p in baseline},**current,str(FIX/'REVIEW_MANIFEST.json'):sha(FIX/'REVIEW_MANIFEST.json')}
write('protected_before_rereview.json',observed)
write('preflight.json',dict(status='PASS',product_current=current,fix_evidence_files=len(m['evidence_files']),previous_evidence_files=len(old['evidence_files']),fix_manifest_last=True,previous_review_unchanged=True,existing_docx=doc,existing_docx_unchanged=True,word_renderer_unchanged=True,changed_runtime=sorted(changed),new_product=[str(ROOT/'tests/test_stage5_s5_r1.py')],reconstructed_source=recovered,changed_functions=scopes,protected_before_count=len(observed),fix_manifest_sha=sha(FIX/'REVIEW_MANIFEST.json')))
print(json.dumps(dict(status='PASS',current_product_files=len(current),fix_evidence_files=len(m['evidence_files']),previous_evidence_files=len(old['evidence_files']),changed_functions=scopes,protected_before_count=len(observed)),ensure_ascii=False))
