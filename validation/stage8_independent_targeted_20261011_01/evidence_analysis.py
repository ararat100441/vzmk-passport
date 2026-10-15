"""Independent package, historical matrix and immutable-version bindings."""
import hashlib,json,xml.etree.ElementTree as E
from pathlib import Path
from zipfile import ZipFile
RUN=Path(__file__).resolve().parent;ROOT=RUN.parents[1];FIX=ROOT/'validation/stage8_fix_20261010_01'
def read(p):return json.loads(p.read_text(encoding='utf-8-sig'))
def save(p,v):p.write_text(json.dumps(v,ensure_ascii=False,indent=2),encoding='utf8')
def sha(p):return hashlib.sha256(Path(p).read_bytes()).hexdigest()
positive=read(RUN/'NATIVE_PRESERVATION.json');before=Path(positive['backup']['path']);after=Path(positive['sandbox'])/'registry/УЗК_TEST.xlsm'
ns={'s':'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
with ZipFile(before) as a,ZipFile(after) as b:
    changed=[n for n in a.namelist() if a.read(n)!=b.read(n)]
    assert set(a.namelist())==set(b.namelist())
    assert a.read('xl/styles.xml')==b.read('xl/styles.xml')
    rels=[n for n in a.namelist() if n.endswith('.rels')]
    assert all(a.read(n)==b.read(n) for n in rels)
    wa=E.fromstring(a.read('xl/workbook.xml'));wb=E.fromstring(b.read('xl/workbook.xml'))
    structure={}
    for tag in ('sheets','definedNames','workbookProtection','externalReferences'):
        nodea=wa.find('s:'+tag,ns);nodeb=wb.find('s:'+tag,ns)
        va=E.tostring(nodea,encoding='unicode') if nodea is not None else None
        vb=E.tostring(nodeb,encoding='unicode') if nodeb is not None else None
        assert va==vb,tag
        structure[tag]=va
    cell_changes={};formulas={}
    for part in a.namelist():
        if not part.startswith('xl/worksheets/') or not part.endswith('.xml'):continue
        ta=E.fromstring(a.read(part));tb=E.fromstring(b.read(part))
        def cells(t):
            return {c.get('r'):{'attributes':c.attrib,'children':[(x.tag,x.text,x.attrib) for x in c]} for c in t.findall('.//s:c',ns)}
        ca,cb=cells(ta),cells(tb)
        assert ca.keys()==cb.keys()
        for coord in ca:
            if ca[coord]!=cb[coord]:cell_changes[part+'!'+coord]={'before':ca[coord],'after':cb[coord]}
        fa={c.get('r'):c.find('s:f',ns).text for c in ta.findall('.//s:c',ns) if c.find('s:f',ns) is not None}
        fb={c.get('r'):c.find('s:f',ns).text for c in tb.findall('.//s:c',ns) if c.find('s:f',ns) is not None}
        assert fa==fb;formulas[part]=fa
    assert set(cell_changes)=={'xl/worksheets/sheet1.xml!C24','xl/worksheets/sheet1.xml!C25','xl/worksheets/sheet2.xml!A1'},cell_changes
save(RUN/'INDEPENDENT_RAW_XLSM.json',{'status':'PASS','before':str(before),'after':str(after),'before_sha256':sha(before),'after_sha256':sha(after),'parts':len(a.namelist()),'changed_parts':changed,'all_relationships_exact':rels,'styles_exact':True,'workbook_structure':structure,'formula_expressions':formulas,'cell_changes':cell_changes})
oldbefore=read(FIX/'PROTECTED_BEFORE.json');oldafter=read(FIX/'final_integrity_02/PROTECTED_AFTER.json');canonical=read(FIX/'CANONICAL_TARGETED_RESULTS.json')
assert oldbefore==oldafter and len(oldbefore)==19618
assert canonical['PASS']==38 and canonical['FAIL']==0 and len(canonical['results'])==38 and all(x['status']=='PASS' for x in canonical['results'])
delivery=read(FIX/'FULL_PRODUCT_BINDING.json');current={p:sha(p) for p in delivery['current']}
assert current==delivery['before_tests']==delivery['after_tests']==delivery['current']
runtime={name:read(FIX/name/'PRODUCT_BEFORE.json') for name in ('targeted_03','targeted_04','targeted_05')}
for name,value in runtime.items():assert value==read(FIX/name/'PRODUCT_AFTER.json')
assert runtime['targeted_03']==runtime['targeted_04']==runtime['targeted_05']
save(RUN/'PRIOR_EVIDENCE_BINDING.json',{'status':'PASS','historical_protected_before_after_exact':True,'historical_protected_count':19618,'current_delivery_matches':True,'delivery_fingerprint':delivery['fingerprint_sha256'],'same_runtime_across_03_04_05':True,'claimed_canonical_38_status':True,'matrix_sources':canonical['sources'],'historical_failed_assertion_preserved':canonical['superseded_failed_check'],'independent_test_results_separate':True})
print('PACKAGE_AND_PRIOR_BINDING_PASS')
