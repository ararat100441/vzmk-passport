"""Independent targeted review. Writes only beside this script; no generators."""
import sys
sys.dont_write_bytecode = True
import json, hashlib, difflib, os
from pathlib import Path
from zipfile import ZipFile
from lxml import etree as E

OUT = Path(__file__).resolve().parent
ROOT = OUT.parents[1]
FIX = ROOT/'validation/stage7_user_e2e_fix_20261009_212419'
OLD = ROOT/'validation/stage7_resume_20261009_175000/ASSEMBLY_01'
PREVIOUS_REVIEW = ROOT/'validation/stage7_independent_review_20261009_01'
FIRST = ROOT.parent/'Гараж казахстан/Паспорт/1_2026 - Документ о качестве.docx'
MANUAL = FIRST.parent/'9 от 06.10.2026/9_2026 - Документ о качестве.docx'
NEW = FIX/'QUALITY_TEST_V2/9_2026 - Документ о качестве (ТЕСТ).docx'
MANIFEST = OLD/'STAGE7_SELECTED_TEST_MANIFEST.json'
RECEIPT = FIX/'QUALITY_TEST_V2/operations/dbf5f02875424ef2a752405a3f90e076/STAGE7_OPERATION_RECEIPT.json'
EXPECTED_NEW = '247f99a361276da5e3ad5db84f1f91ab5dd9d5527e21a6ae2b4615c3aae6ac91'
NS = {'w':'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
def read(p): return json.loads(Path(p).read_text(encoding='utf-8-sig'))
def save(n, x): (OUT/n).write_text(json.dumps(x, ensure_ascii=False, indent=2, default=str), encoding='utf8')
def sha(p):
    p=Path(p)
    assert p.name.casefold() not in ('взмк.xlsx','узк.xlsм','узк.xlsm'), str(p)
    return hashlib.sha256(p.read_bytes()).hexdigest()
def actual(p): return sha(p) if Path(p).is_file() else None
def xml(p):
    with ZipFile(p) as z: return E.fromstring(z.read('word/document.xml'))
def text(x): return ''.join(x.xpath('.//w:t/text()',namespaces=NS))
def c14(x): return E.tostring(x,method='c14n') if x is not None else None
def cells(x): return [[r.findall('w:tc',NS) for r in t.findall('w:tr',NS)] for t in x.findall('w:body/w:tbl',NS)]
def mismatches(d): return [{'path':p,'expected':h,'actual':actual(p)} for p,h in d.items() if actual(p)!=h]

def hash_review():
    protected=read(FIX/'PROTECTED_BEFORE.json')
    product=read(FIX/'PRODUCT_AFTER_FINAL_TESTS.json')
    oldproduct=read(FIX/'PRODUCT_BEFORE.json')
    audit=read(ROOT/'validation/stage7_quality_audit_20261009/source_before.json')
    previous=read(PREVIOUS_REVIEW/'SOURCE_AFTER.json')
    em=read(FIX/'EVIDENCE_MANIFEST.json')
    evidence={p:v['sha256'] for p,v in em['evidence'].items()}
    # Scope is the explicitly protected set, current code, and the supplied fix evidence.
    scope=protected | product | evidence
    scope[str(FIX/'EVIDENCE_MANIFEST.json')]=sha(FIX/'EVIDENCE_MANIFEST.json')
    for p in (FIRST,MANUAL,NEW,MANIFEST,RECEIPT,OLD/'QUALITY_TEST/9_2026 - Документ о качестве (ТЕСТ).docx'):
        scope[str(p)]=sha(p)
    snapshot={p:actual(p) for p in scope}
    save('SOURCE_BEFORE.json',snapshot)
    changes=[]
    for p in sorted(set(oldproduct)|set(product)):
        if oldproduct.get(p)!=product.get(p):
            changes.append({'path':p,'before':oldproduct.get(p),'after':product.get(p),'kind':'ADDED' if p not in oldproduct else 'MODIFIED'})
    diff=[]; archive_checks=[]
    for name in ('word.py','core.py','provenance.py'):
        a=FIX/'code_before'/name;b=ROOT/'stage7_quality'/name
        prior=previous[str(b)]['sha256']
        archive_checks.append({'name':name,'archived_sha256':sha(a),'previous_independent_review_sha256':prior,'equal':sha(a)==prior==oldproduct[str(b)]})
        diff.extend(difflib.unified_diff(a.read_text(encoding='utf8').splitlines(True),b.read_text(encoding='utf8').splitlines(True),fromfile='accepted_before/stage7_quality/'+name,tofile='current/stage7_quality/'+name))
    supplied_diff_equal=''.join(diff)==(FIX/'EXACT_PRODUCT_DIFF.diff').read_text(encoding='utf8')
    newhelper=ROOT/'stage7_quality/presentation_revision.py'
    diff.extend(difflib.unified_diff([],newhelper.read_text(encoding='utf8').splitlines(True),fromfile='/dev/null',tofile='current/stage7_quality/presentation_revision.py'))
    (OUT/'INDEPENDENT_EXACT_PRODUCT_DIFF.diff').write_text(''.join(diff),encoding='utf8')
    refs={str(p):{'actual_sha256':sha(p),'stage7_audit_sha256':audit[str(p)]['sha256'],'equal':sha(p)==audit[str(p)]['sha256']} for p in (FIRST,MANUAL)}
    hashes={k:{'path':str(p),'sha256':sha(p),'bytes':p.stat().st_size} for k,p in {'new':NEW,'first':FIRST,'manual':MANUAL,'manifest':MANIFEST,'receipt':RECEIPT,'old':OLD/'QUALITY_TEST/9_2026 - Документ о качестве (ТЕСТ).docx'}.items()}
    # Enumerate the bound product directories independently to catch extra source files.
    parentdirs={Path(p).parent for p in product}
    candidates=set()
    for parent in parentdirs:
        for p in parent.iterdir():
            if p.is_file() and (p.suffix.lower() in ('.py','.ps1','.cmd') or p.name in ('accepted_checkpoint.json','CATALOG_SCHEMA_STAGE6_CONTRACT.txt')):
                candidates.add(str(p))
    unlisted=sorted(candidates-set(product))
    result={'protected_count':len(protected),'protected_before_after_equal':protected==read(FIX/'PROTECTED_AFTER.json'),
        'protected_current_mismatches':mismatches(protected),'product_count':len(product),'product_current_mismatches':mismatches(product),
        'evidence_count':len(evidence),'evidence_current_mismatches':mismatches(evidence),'scope_count':len(snapshot),
        'new_sha_expected':hashes['new']['sha256']==EXPECTED_NEW,'new_distinct_from_old':hashes['new']['path']!=hashes['old']['path'] and hashes['new']['sha256']!=hashes['old']['sha256'],
        'old_sha_matches_previous_review':hashes['old']['sha256']==previous[hashes['old']['path']]['sha256'],
        'old_manifest_matches_previous_review':hashes['manifest']['sha256']==previous[str(MANIFEST)]['sha256'],
        'references':refs,'product_changes':changes,'extra_product_source_files':unlisted,'before_archives':archive_checks,'supplied_diff_equal':supplied_diff_equal,
        'stage1_6_changes':[c for c in changes if Path(c['path']).parent!=ROOT/'stage7_quality'],
        'production_access':False}
    save('ACTUAL_HASHES.json',hashes);save('HASH_REVIEW.json',result)
    print(json.dumps(result,ensure_ascii=False,indent=2))

def formatting(node):
    # Ignore revision-session identifiers only; retain ALL paragraph and run properties.
    def sem(n):
        if n is None:return None
        return (E.QName(n).localname,tuple(sorted((E.QName(k).localname,v) for k,v in n.attrib.items() if not E.QName(k).localname.startswith('rsid'))),tuple(sem(c) for c in n))
    return [{'text':text(p),'pPr':sem(p.find('w:pPr',NS)),
        'runs':[{'text':text(r),'rPr':sem(r.find('w:rPr',NS))} for r in p.findall('w:r',NS)]} for p in node.findall('w:p',NS)]

def structure_review():
    q=xml(NEW);f=xml(FIRST);m=xml(MANUAL);old=xml(OLD/'QUALITY_TEST/9_2026 - Документ о качестве (ТЕСТ).docx')
    qc,fc,mc,oc=map(cells,(q,f,m,old))
    fr=f.findall('w:body/w:tbl',NS)[0].findall('w:tr',NS)
    qr=q.findall('w:body/w:tbl',NS)[0].findall('w:tr',NS)
    start=next(i for i,r in enumerate(fr) if 'Согласно условиям договора на поставку конструкции защищены от коррозии:' in text(r))
    end=next(i for i,r in enumerate(fr[start+1:],start+1) if 'Документ о качестве составлен на основании приемочных актов:' in text(r))
    same=[c14(r)==c14(s) for r,s in zip(fr[start:end],qr[start:end])]
    for label,rs in [('FIRST',fr),('NEW',qr)]: (OUT/f'POINT12_{label}_CANONICAL.xml').write_bytes(b'<point12>'+b''.join(c14(r) for r in rs[start:end])+b'</point12>')
    def parts(a,b):
        with ZipFile(a) as x,ZipFile(b) as y:
            return {'entry_names_equal':x.namelist()==y.namelist(),'changed_parts':[n for n in x.namelist() if x.read(n)!=y.read(n)]}
    for label,x in [('NEW',q),('FIRST',f),('MANUAL',m)]: (OUT/(label+'_TEXT.txt')).write_text('\n'.join(text(p) for p in x.xpath('.//w:p',namespaces=NS)),encoding='utf8')
    actual_cells=[]
    for ti,t in enumerate(oc):
        for ri,row in enumerate(t):
            for ci,c in enumerate(row):
                if c14(c)!=c14(qc[ti][ri][ci]):actual_cells.append([ti,ri,ci])
    result={'tables':len(qc),'rows':[len(t) for t in qc],'sections':len(q.findall('.//w:sectPr',NS)),
        'point12_rows_1based':[start+1,end],'point12_canonical_xml_equal':all(same),'point12_text':[text(r) for r in qr[start:end]],
        'first_package':parts(FIRST,NEW),'previous_package':parts(OLD/'QUALITY_TEST/9_2026 - Документ о качестве (ТЕСТ).docx',NEW),
        'point13':{'actual':formatting(qc[0][50][0]),'manual':formatting(mc[0][50][0])},
        'point14':{'actual':formatting(qc[0][53][0]),'manual':formatting(mc[0][53][0])},
        'point14_explanation_equal':c14(qr[54])==c14(fr[54]),'point14_explanation':text(qr[54]),
        'contract':{'label':text(qc[0][11][2]),'value':text(qc[0][11][3]),'manual_label':text(mc[0][11][2]),'manual_value':text(mc[0][11][3])},
        'changed_cells_vs_previous':actual_cells,
        'signature_table_equal_first':c14(f.findall('w:body/w:tbl',NS)[1])==c14(q.findall('w:body/w:tbl',NS)[1]),
        'signature_cells_equal_old':all(c14(c)==c14(qc[1][ri][ci]) for ri,row in enumerate(oc[1]) for ci,c in enumerate(row)),
        'drawings':len(q.xpath('.//w:drawing|.//w:pict',namespaces=NS)),
        'first_drawings':len(f.xpath('.//w:drawing|.//w:pict',namespaces=NS))}
    for xp in ('.//w:numPr','.//w:sectPr','.//w:tblPr','.//w:trPr','.//w:instrText','.//w:fldChar'):
        result[xp+'_equal_first']=[c14(n) for n in f.xpath(xp,namespaces=NS)]==[c14(n) for n in q.xpath(xp,namespaces=NS)]
    save('OOXML_REVIEW.json',result)
    print(json.dumps({k:v for k,v in result.items() if k not in ('point13','point14')},ensure_ascii=False,indent=2))
    for label in ('point13','point14'):
        a=result[label]['actual'];b=result[label]['manual']
        print(label,'text', [p['text'] for p in a]==[p['text'] for p in b], 'all_raw_runs_equal',a==b)

def after():
    before=read(OUT/'SOURCE_BEFORE.json');current={p:actual(p) for p in before};save('SOURCE_AFTER.json',current)
    changed=[{'path':p,'before':h,'after':current[p]} for p,h in before.items() if current[p]!=h]
    product=read(FIX/'PRODUCT_AFTER_FINAL_TESTS.json')
    result={'status':'PASS' if not changed else 'FAIL','scope_count':len(before),'changes':changed,
        'protected_count':len(read(FIX/'PROTECTED_BEFORE.json')),'protected_current_mismatches':mismatches(read(FIX/'PROTECTED_BEFORE.json')),
        'product_count':len(product),'product_current_mismatches':mismatches(product),
        'source_writes':0,'document_generations':{'VIK':0,'UZK':0,'AKZ':0,'QUALITY':0},
        'official_quality_guard_test_invocations':1,'production_access':False}
    save('SOURCE_INTEGRITY.json',result);print(json.dumps(result,ensure_ascii=False,indent=2))

if __name__=='__main__':globals()[sys.argv[1]]()
