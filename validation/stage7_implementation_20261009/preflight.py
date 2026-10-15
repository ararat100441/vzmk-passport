import sys,json,hashlib,os,zipfile
from pathlib import Path
from collections import Counter
from decimal import Decimal
import openpyxl
OUT=Path(__file__).resolve().parent;ROOT=OUT.parents[1];sys.path.insert(0,str(ROOT))
from passport_app.registry import read_registry,select_project
from uzk_stage5.inputs import visible_vom
def sha(p):return hashlib.sha256(p.read_bytes()).hexdigest()
def save(n,d):(OUT/n).write_text(json.dumps(d,ensure_ascii=False,indent=2,default=str),encoding='utf8')
def load(p):return json.loads(p.read_text(encoding='utf-8-sig'))
def walk(root):
    for b,dirs,files in os.walk(root,followlinks=False):
        dirs[:]=[d for d in dirs if d not in ('validation','.codebase-index','node_modules','__pycache__') and not (Path(b)/d).is_symlink() and not (Path(b)/d).is_junction()]
        for n in files:
            p=Path(b)/n
            if not n.startswith('~$') and p.suffix.lower()!='.pdf':yield p
P=ROOT.parent/'Гараж казахстан/Паспорт'
sources=list(walk(ROOT))+list(walk(P))
save('SOURCE_BEFORE.json',{str(p):sha(p) for p in sources})
registry=read_registry(ROOT/'DEV/ВЗМК_DEV.xlsx');project=select_project(registry,'Гараж Казахстан')
selected=[e for e in registry.projects[project.name] if e['number']=='9/2026'];assert len(selected)==1
assert registry.digest=='8e56f68518c70cade9efb09084dd67cd5d72e07e3fdc813b8515030e9cb98e80'
template=P/'1_2026 - Документ о качестве.docx';assert sha(template)=='4df7234d37030188030142ed6a519dc2a923aada4f90e58a7b5140c0fd54f9cc'
vom=visible_vom(P/'9 от 06.10.2026/ВОМ отгрузочный счет 183  (5).xls');save('VOM_CURRENT.json',vom)
assert (vom['visible_count'],len(vom['hidden_data_rows']),sum(r['quantity'] for r in vom['rows']),vom['mass_kg'])==(104,397,288,'15822.14')
vm=Counter((r['mark'],r['quantity'],Decimal(r['mass_kg']).quantize(Decimal('.01'))) for r in vom['rows'])
def aggregate(records):
    result={}
    for mark,qty,mass in records:
        entry=result.setdefault(mark,{'quantity':0,'mass':Decimal(0)})
        entry['quantity']+=qty;entry['mass']+=mass
    return result
vm_aggregate=aggregate(vm.elements())
journal=load(ROOT/'e2e_user_stage1/.pasport-transactions/26a5854c50dd4bf694081e2e0e5ab938/journal.json')
appendices=[P/'9 от 06.10.2026/9_2026 - Приложение № 1.xlsx',Path(journal['passport'])/'9_2026 - Приложение № 1.xlsx']
apps=[]
for path in appendices:
    book=openpyxl.load_workbook(path,data_only=True,read_only=True);s=book.worksheets[0];body=[]
    for row in s.iter_rows(min_row=5,values_only=True):
        if not row[0]:continue
        if str(row[0]).strip().lower()=='итого':total=row;break
        body.append({'mark':row[0],'quantity':row[2],'mass_unit':row[3],'mass_total':row[4]})
    actual=Counter((r['mark'],int(r['quantity'] or 0),Decimal(str(r['mass_total'] or 0)).quantize(Decimal('.01'))) for r in body)
    aa=aggregate(actual.elements())
    record={'path':str(path),'sha256':sha(path),'header':s['A1'].value,'object':s['A2'].value,'body_count':len(body),'body':body,'total':total,'matches_visible_vom':actual==vm,'matches_visible_vom_per_mark':aa==vm_aggregate,'mark_quantity_mass_differences':{m:{'appendix':aa.get(m),'vom':vm_aggregate.get(m)} for m in set(aa)|set(vm_aggregate) if aa.get(m)!=vm_aggregate.get(m)},'extra_records':list((actual-vm).elements()),'missing_records':list((vm-actual).elements()),'stage1_journal_sha':journal['appendix_hash'],'matches_stage1_journal_sha':sha(path)==journal['appendix_hash']}
    apps.append(record);book.close()
save('APPENDIX_CANDIDATES.json',apps)
sample=load(ROOT/'validation/stage5_sample_results_20261008_130441/new_document.json');sp=Path(sample['output']);assert sp.exists() and sha(sp)==sample['sha256']=='8d3a2769baf6d381a0675804ef7f0806896f1b611e1498ea58677d3fdac1fd1b'
ak=ROOT/'validation/stage6_fix_r1_r2_20261009_094707/NEW_TEST_9_2026';pair=load(ak/'PAIR.json')
assert pair['sha256']['ТЕСТ АКЗ №0.xlsx']=='b640796570a8d04b6bca404205d99eae4c7005c700587812ed25bc99e715eb1a'
assert pair['sha256']['ТЕСТ Приложение к АКЗ №0.xlsx']=='7766c92df82df2484eb82cd6917769f5d9ede912b3f0e8f26e3f8aaabdd357f2'
for n,h in pair['sha256'].items():assert sha(ak/n)==h
save('INPUT_PREFLIGHT.json',{'registry_sha256':registry.digest,'row':selected[0]['row'],'template_sha256':sha(template),'vom_sha256':sha(Path(vom['path'])),'sample':sample,'akz_pair':pair,'akz_folder':str(ak),'appendix_match_count':sum(a['matches_visible_vom_per_mark'] for a in apps)})
print(json.dumps({'registry_row':selected[0]['row'],'vom':{'visible':vom['visible_count'],'hidden':len(vom['hidden_data_rows']),'mass':vom['mass_kg']},'appendices':[{k:a[k] for k in ('path','body_count','matches_visible_vom_per_mark','matches_stage1_journal_sha','total','mark_quantity_mass_differences')} for a in apps]},ensure_ascii=False,default=str,indent=2))
