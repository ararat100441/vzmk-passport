from pathlib import Path
import hashlib,json,sys
from openpyxl import load_workbook
sys.path.insert(0,str(Path(__file__).resolve().parents[1]))
from passport_app.registry import read_registry
base=Path(r'C:\Users\arara\Desktop\ВЗМК рабочий стол\Работа')
out=Path(__file__).parent/'production_100';out.mkdir(exist_ok=True)
registry=base/'ВЗМК (2).xlsx'
sources=[registry,base/'УЗК.xlsm']
for folder in (base/'Документы/ТЭМПО',base/'Документы/Гараж казахстан',base/'Шаблоны'):
    sources += [p for p in folder.rglob('*') if p.suffix.lower() in ('.xlsx','.xlsm','.docx') and not p.name.startswith('~$')]
old=base/'Документы/Pasport/stage9_dist_0916/VZMK-Passport-0.9.16-DEVTEST/VZMK-Passport-0.9.16-DEVTEST.exe'
sources.append(old)
hashes={str(p):hashlib.sha256(p.read_bytes()).hexdigest() for p in sources if p.is_file()}
assert not (out/'SOURCE_BEFORE.json').exists()
(out/'SOURCE_BEFORE.json').write_text(json.dumps(hashes,ensure_ascii=False,indent=2),encoding='utf8')
r=read_registry(registry)
w=load_workbook(registry,data_only=True)
data={}
for project in ('ТЭМПО 1','ТЭМПО 2','Гараж Казахстан'):
    entries=r.projects.get(project,[])
    data[project]=[{'row':e['row'],'number':e['number'],'mass':e['mass'],'values':[c.value for c in w[r.sheet][e['row']]][:18]} for e in entries]
w.close()
first=[p for p in sources if 'ТЭМПО' in str(p) and p.suffix.lower() in ('.xlsx','.docx') and ('1_1_2026' in p.name or ('1 от ' in str(p.parent) and '1 выдача' in str(p)))]
voms=[p for p in (base/'Документы/ТЭМПО/1 выдача').rglob('*') if p.is_file() and ('ВОМ' in p.name.upper() or '1_1_2026' in p.name)]
result={'registry':str(registry),'registry_sha256':r.digest,'projects':data,'first_templates':[str(p) for p in first], 'voms':[str(p) for p in voms],'protected_files':len(hashes)}
(out/'INVENTORY.json').write_text(json.dumps(result,ensure_ascii=False,indent=2,default=str),encoding='utf8')
print(json.dumps({'projects':{k:[(e['row'],e['number'],e['mass']) for e in v] for k,v in data.items()},'first_templates':[str(p.relative_to(base)) for p in first],'voms':[str(p.relative_to(base)) for p in voms][:20],'protected':len(hashes)},ensure_ascii=False,default=str))
