"""Read-only targeted passport-source search and two-row registry comparison."""
from pathlib import Path
from decimal import Decimal
import json,hashlib,sys,re,subprocess
from openpyxl import load_workbook
import xlrd
ROOT=Path(__file__).resolve().parents[1];sys.path.insert(0,str(ROOT))
from passport_app.vom import read_vom
from production_app.storage import sha,save,read

out=ROOT/'validation/production_100/acceptance_sources_final';out.mkdir(exist_ok=False)
work=Path(r'C:\Users\arara\Desktop\ВЗМК рабочий стол\Работа')
original=work/'Документы/Pasport';actual=work/'ВЗМК (2).xlsx'
dev=original/'DEV/ВЗМК_DEV.xlsx'
previous=read(ROOT/'validation/production_100/SOURCE_BEFORE.json')
protected={p:sha(p) for p in previous if Path(p).is_file()}
protected[str(dev)]=sha(dev)
for dist in ('stage9_dist_0916','stage9_dist_100_rc1'):
    protected.update({str(p):sha(p) for p in (original/dist).rglob('*') if p.is_file()})
status=subprocess.check_output(['git','status','--porcelain=v1','-z'],cwd=original)
(out/'ORIGINAL_GIT_BEFORE.bin').write_bytes(status)
untracked=subprocess.check_output(['git','ls-files','--others','--exclude-standard','-z'],cwd=original)
for p in untracked.split(b'\0'):
    if p:protected[str(original/p.decode('utf8'))]=sha(original/p.decode('utf8'))
save(out/'PROTECTED_BEFORE.json',protected)
conflicts={}
for label,path in (('working',actual),('DEV',dev)):
    w=load_workbook(path,data_only=True,keep_links=True)
    try:
        s=w.worksheets[0]
        conflicts[label]={'path':str(path),'sha256':sha(path),'rows':[
            {'row':row,'project':s.cell(row,1).value,'number':s.cell(row,2).value,'contract':s.cell(row,3).value,
             'customer':s.cell(row,4).value,'object':s.cell(row,5).value,'cipher':s.cell(row,6).value,'mass_kg':s.cell(row,7).value}
            for row in (342,353,401,407) if row<=s.max_row]}
    finally:w.close()
save(out/'CONFLICTS_401_407.json',conflicts)
appendix=work/'Документы/ТЭМПО/1 выдача/Паспорта/1 от 02.09.2026/1_1_2026 - Приложение № 1.xlsx'
expected=read_vom(appendix);marks={str(r[0]) for r in expected.rows}
candidates=set()
for folder in (work/'Документы/ТЭМПО/1 выдача',work/'Приход',work/'Приход (2)'):
    candidates.update(p for p in folder.rglob('*') if p.suffix.lower() in ('.xls','.xlsx','.xlsm') and p.is_file() and 'Паспорта' not in p.parts and not p.name.startswith('~$'))
downloads=Path.home()/'Downloads'
candidates.update(p for p in downloads.rglob('*') if p.is_file() and p.suffix.lower() in ('.xls','.xlsx','.xlsm')
    and re.search(r'ВОМ|ТЭМПО|УПАКОВ|ОТГРУЗ',p.name,re.I) and not p.name.startswith('~$'))
evidence=[];errors=[];mass_matches=[];matches=[]
for path in sorted(candidates):
    initial=sha(path);entry={'path':str(path),'sha256':initial,'sheets':[]}
    try:
        if path.suffix.lower()=='.xls':
            w=xlrd.open_workbook(path,on_demand=True)
            def xls_rows(sheet):
                for rn in range(min(sheet.nrows,20000)):yield sheet.row_values(rn,0,min(sheet.ncols,60))
            sheets=[(s.name,xls_rows(s)) for s in w.sheets()]
        else:
            w=load_workbook(path,read_only=True,data_only=True,keep_links=False)
            sheets=[(s.title,s.iter_rows(max_row=min(s.max_row or 20000,20000),max_col=min(s.max_column or 60,60),values_only=True)) for s in w.worksheets]
        try:
            for title,rows in sheets:
                found=set();mass_found=False;hits=[]
                for rn,row in enumerate(rows,1):
                    for cn,value in enumerate(row,1):
                        if str(value) in marks:found.add(str(value));hits.append((rn,cn,str(value)))
                        if isinstance(value,(int,float)) and not isinstance(value,bool) and abs(Decimal(str(value))-expected.mass)<Decimal('.005'):mass_found=True
                item={'sheet':title,'expected_marks_found':len(found),'mass_14747_36_found':mass_found}
                entry['sheets'].append(item)
                if mass_found:mass_matches.append({'path':str(path),**item})
                if found==marks:
                    matches.append({'path':str(path),'sha256':initial,**item,'mark_hits':hits})
        finally:
            if path.suffix.lower()=='.xls':w.release_resources()
            else:w.close()
        if sha(path)!=initial:raise RuntimeError('SOURCE_CHANGED_DURING_READ')
        evidence.append(entry)
    except Exception as exc:errors.append({'path':str(path),'error':str(exc),'sha256':initial})
result={'passport':'1/1/2026','expected_marks':len(expected.rows),'expected_quantity':str(expected.quantity),'expected_mass_kg':str(expected.mass),
    'reference_appendix':str(appendix),'searched_candidates':len(candidates),'readable':len(evidence),'errors':errors,'mass_matches':mass_matches,'full_mark_matches':matches,
    'confirmed_vom':None,'status':'STOP_PASSPORT_SCENARIO_UNTIL_VOM_PROVEN','scope':'Targeted project/receipt folders and recursively named VOM/TEMPO/packing/shipping workbooks in Downloads'}
save(out/'VOM_SEARCH.json',result);save(out/'VOM_SEARCH_READS.json',evidence)
print(json.dumps({'protected':len(protected),'conflicts':conflicts,'candidates':len(candidates),'readable':len(evidence),'errors':len(errors),'mass_matches':mass_matches,'full_mark_matches':matches},ensure_ascii=False,indent=2))
