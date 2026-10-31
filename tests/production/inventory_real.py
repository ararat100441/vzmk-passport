"""Read-only inventory of local production sources; never changes an input."""
import argparse,json,sys,zipfile,re,hashlib
from pathlib import Path
from xml.etree import ElementTree as ET
ROOT=Path(__file__).resolve().parents[2];sys.path.insert(0,str(ROOT))
from production_app import model
from production_app.storage import sha

def main():
    p=argparse.ArgumentParser();p.add_argument('--output',required=True);a=p.parse_args()
    work=Path(r'C:\Users\PUBLIC_USER\Desktop\ВЗМК рабочий стол\Работа')
    documents=work/'Документы';registry=work/'ВЗМК (2).xlsx'
    roots=[documents/'ТЭМПО/1 выдача/Паспорта',documents/'Гараж казахстан/Паспорт']
    inputs=[registry]+[f for root in roots for f in root.rglob('*') if f.is_file() and not f.name.startswith('~$')]
    before={str(f):sha(f) for f in inputs};r=model.read_registry(registry)
    projects={}
    for name in ('ТЭМПО 1','Гараж Казахстан'):
        try:projects[name]=vars(model.select(r,name))
        except Exception as e:projects[name]={'error':str(e)}
    contracts=[]
    for f in inputs:
        if f.suffix.lower()=='.docx':
            with zipfile.ZipFile(f) as z:
                x=ET.fromstring(z.read('word/document.xml'))
                ns={'w':'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
                text=''.join(t.text or '' for t in x.findall('.//w:t',ns))
                tokens=re.findall(r'\{\{[A-Z_0-9]+\}\}',text)
                tags=[e.attrib for e in x.findall('.//w:tag',ns)]
            contracts.append({'path':str(f),'tokens':tokens,'content_controls':tags,'text_prefix':text[:180]})
    voms=[]
    for root in roots:
        for f in root.rglob('*.xlsm'):
            try:
                v=model.visible_vom(f)
                matches=[{'project':name,'number':e['number'],'row':e['row']} for name,es in r.projects.items() for e in es if name in projects and e['mass'] is not None and abs(float(v.mass)-float(e['mass']))<.005]
                voms.append({'path':str(f),'marks':len(v.rows),'quantity':str(v.quantity),'mass':str(v.mass),'registry_matches_by_mass_only':matches})
            except Exception as e:voms.append({'path':str(f),'error':str(e)})
    result={'registry':str(registry),'registry_sha256':before[str(registry)],'projects':projects,'files':before,'word_contracts':contracts,'voms':voms,
            'catalogs':[str(f) for root in roots for f in root.parent.rglob('*Справочник*марок*.xlsx')],
            'sources_unchanged':all(sha(f)==h for f,h in before.items())}
    target=Path(a.output);target.parent.mkdir(parents=True,exist_ok=True);target.write_text(json.dumps(result,ensure_ascii=False,indent=2,default=str),encoding='utf8')
    print(json.dumps({k:result[k] for k in ('registry','registry_sha256','projects','catalogs','sources_unchanged')},ensure_ascii=False,default=str))
    print('files',len(inputs),'Word',len(contracts),'VOM',len(voms))
if __name__=='__main__':main()
