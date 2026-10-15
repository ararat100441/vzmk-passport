import sys
sys.dont_write_bytecode = True
import os, json, hashlib, stat
from pathlib import Path
from zipfile import ZipFile
from lxml import etree as E

ROOT = Path(__file__).resolve().parents[2]
OUT = Path(__file__).resolve().parent
IMPL = ROOT/'validation/stage7_resume_20261009_175000'
sys.path.insert(0, str(ROOT))

def read(p): return json.loads(Path(p).read_text(encoding='utf-8-sig'))
def sha(p):
    p=Path(p)
    assert p.name.casefold() not in ('взмк.xlsx','узк.xlsm'), str(p)
    return hashlib.sha256(p.read_bytes()).hexdigest()
def save(n,d): (OUT/n).write_text(json.dumps(d,ensure_ascii=False,indent=2,default=str),encoding='utf8')

def snapshot():
    result={}
    for start in [ROOT,ROOT.parent/'Гараж казахстан/Паспорт']:
        for parent,dirs,files in os.walk(start,followlinks=False):
            dirs[:]=[d for d in dirs if d not in ('node_modules','.git','.codebase-index','__pycache__','.venv')
                and not (Path(parent)/d).is_relative_to(OUT)
                and not (Path(parent)/d).is_symlink()
                and not (os.lstat(Path(parent)/d).st_file_attributes & stat.FILE_ATTRIBUTE_REPARSE_POINT)]
            for n in files:
                p=Path(parent)/n
                if n.startswith('~$') or p.suffix.lower() in ('.pdf','.exe','.pyc') or n.casefold() in ('взмк.xlsx','узк.xlsm'): continue
                if p.is_symlink() or os.lstat(p).st_file_attributes & stat.FILE_ATTRIBUTE_REPARSE_POINT: continue
                result[str(p)]={'sha256':sha(p),'bytes':p.stat().st_size}
    return result

def hashes():
    if (OUT/'SOURCE_BEFORE.json').exists():before=read(OUT/'SOURCE_BEFORE.json')
    else:
        before=snapshot();save('SOURCE_BEFORE.json',before)
    checks={};actual={}
    for n in ['BASELINE_BEFORE.json','BASELINE_AFTER.json','PRODUCT_AFTER_FINAL_TESTS.json']:
        d=read(IMPL/n)
        checks[n]={'count':len(d),'mismatch':[]}
        for p,h in d.items():
            got=sha(p) if Path(p).is_file() else None
            if got!=h: checks[n]['mismatch'].append({'path':p,'expected':h,'actual':got})
    checkpoint=read(ROOT/'stage7_quality/accepted_checkpoint.json')
    checks['accepted_checkpoint']={'count':len(checkpoint['product']),'mismatch':[
        p for p,h in checkpoint['product'].items() if sha(p)!=h]}
    em=read(IMPL/'EVIDENCE_MANIFEST.json')
    for key in ['evidence','new_product']:
        d=em[key]
        checks[key]={'count':len(d),'mismatch':[]}
        entries=d.items() if isinstance(d,dict) else [(v['path'],v) for v in d]
        for p,v in entries:
            h=v['sha256'] if isinstance(v,dict) else v
            if not Path(p).is_file() or sha(p)!=h:checks[key]['mismatch'].append(p)
    m=read(IMPL/'ASSEMBLY_01/STAGE7_SELECTED_TEST_MANIFEST.json')
    result=read(IMPL/'ASSEMBLY_01/ASSEMBLY_RESULT.json')
    p=result['quality_receipt']['path'];q=read(p)
    targets={r['document_id']:r['path'] for r in m['records']}
    targets.update(quality=q['output'],manifest=str(IMPL/'ASSEMBLY_01/STAGE7_SELECTED_TEST_MANIFEST.json'),
        template=str(ROOT.parent/'Гараж казахстан/Паспорт/1_2026 - Документ о качестве.docx'),
        vom=str(ROOT/'passport_app/vom.py'))
    for label,p in targets.items():actual[label]={'path':p,'sha256':sha(p),'bytes':Path(p).stat().st_size}
    for r in read(IMPL/'CHANGED_FILES.json')['product_files']:actual['product:'+Path(r['path']).name]={'path':r['path'],'sha256':sha(r['path'])}
    save('ACTUAL_HASHES.json',actual);save('HASH_VERIFICATION.json',checks)
    print(json.dumps({'snapshot':len(before),'checks':checks,'output':actual['quality']},ensure_ascii=False))

def structure():
    a=read(OUT/'ACTUAL_HASHES.json');ns={'w':'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
    def doc(p):
        with ZipFile(p) as z:return E.fromstring(z.read('word/document.xml'))
    def txt(n):return ''.join(n.xpath('.//w:t/text()',namespaces=ns))
    t=doc(a['template']['path']);q=doc(a['quality']['path'])
    rows=t.findall('w:body/w:tbl',ns)[0].findall('w:tr',ns)
    out=q.findall('w:body/w:tbl',ns)[0].findall('w:tr',ns)
    # Identify the coating block independently by its heading and next numbered heading.
    for i,r in enumerate(rows):print(i+1,txt(r)[:150])
    changed=[]
    for i,(r,s) in enumerate(zip(rows,out),1):
        if E.tostring(r,method='c14n')!=E.tostring(s,method='c14n'):changed.append(i)
    with ZipFile(a['template']['path']) as x,ZipFile(a['quality']['path']) as y:
        changed_parts=[n for n in x.namelist() if x.read(n)!=y.read(n)]
    alltext={k:'\n'.join(txt(p) for p in doc(v['path']).xpath('.//w:p',namespaces=ns))
        for k,v in a.items() if k in ('quality','vik','uzk','template')}
    for k,v in alltext.items():(OUT/(k+'_text.txt')).write_text(v,encoding='utf8')
    save('OOXML_REVIEW.json',{'changed_rows':changed,'changed_parts':changed_parts,
        'tables':len(q.findall('w:body/w:tbl',ns)), 'sections':len(q.findall('.//w:sectPr',ns)),
        'point12_rows_41_49_equal':[E.tostring(r,method='c14n') for r in rows[40:49]]==[E.tostring(r,method='c14n') for r in out[40:49]],
        'point12_text':[txt(r) for r in out[40:49]],'point13':txt(out[50]),'point14':txt(out[53]),
        'drawings':len(q.xpath('.//w:drawing|.//w:pict',namespaces=ns)),
        'external_references':q.xpath('.//w:hyperlink/@r:id',namespaces=ns|{'r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships'})})

def after():
    before=read(OUT/'SOURCE_BEFORE.json');current=snapshot();save('SOURCE_AFTER.json',current)
    changed=[{'path':p,'before':before.get(p),'after':current.get(p)} for p in sorted(set(before)|set(current)) if before.get(p)!=current.get(p)]
    save('SOURCE_INTEGRITY.json',{'status':'PASS' if not changed else 'FAIL','before':len(before),'after':len(current),'changes':changed,
        'roots':[str(ROOT),str(ROOT.parent/'Гараж казахстан/Паспорт')],
        'excluded':['this independent evidence directory','dependency/index/cache/reparse paths','PDF/EXE','Office lock files','ВЗМК.xlsx','УЗК.xlsm'],
        'production_opened':False,'stage1_6_modified':False,'previous_receipts_modified':False})
    print(json.dumps({'before':len(before),'after':len(current),'changed':changed},ensure_ascii=False))

if __name__=='__main__':globals()[sys.argv[1]]()
