from pathlib import Path
import hashlib
OUT=Path(__file__).resolve().parent;ROOT=OUT.parents[1]
def sha(p):return hashlib.sha256(Path(p).read_bytes()).hexdigest()
def save(p,d):
    with Path(p).open('x',encoding='utf8') as f:json.dump(d,f,ensure_ascii=False,indent=2,default=str)
from openpyxl import load_workbook
import json
baseline=json.loads((OUT/'TEMPLATE_BINDINGS_BEFORE.json').read_text(encoding='utf8'))
source_root=ROOT.parent/'Гараж казахстан/Паспорт'
rows=[]
for ref in baseline['pairs']:
    p=source_root/Path(ref['path']).name
    if not p.exists():rows.append({'path':str(p),'status':'MISSING'});continue
    current=sha(p);b=load_workbook(p,read_only=True,data_only=False,keep_links=False)
    s=b['4'] if ref['kind']=='act' else b.active
    cells={x:str(s[x].value) for x in ('E4','H4','E5','E6','E7','B44','A1','A2')}
    b.close();assert sha(p)==current
    rows.append({'path':str(p),'variant':ref['variant'],'kind':ref['kind'],'sha256':current,'same_as_sealed_dev':current==ref['sha256'],'cells':cells,'read_only':True})
save(OUT/'CURRENT_PROJECT_TEMPLATES.json',{'registry_source':'sealed DEV only','first_passport':baseline['project'],'template_reads_only':True,'files':rows})
print(json.dumps(rows,ensure_ascii=False,indent=2))
