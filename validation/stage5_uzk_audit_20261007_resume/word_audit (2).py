import json,re,hashlib,sys
from pathlib import Path
from zipfile import ZipFile
from xml.etree import ElementTree as ET
sys.dont_write_bytecode=True
OUT=Path(__file__).resolve().parent;APP=OUT.parent.parent;ROOT=APP.parent
sys.path.insert(0,str(APP))
from passport_app.registry import read_registry,select_project
from passport_app.catalog import key
NS={'w':'http://schemas.openxmlformats.org/wordprocessingml/2006/main','a':'http://schemas.openxmlformats.org/drawingml/2006/main'}
W='{'+NS['w']+'}'
def txt(el):return ''.join(t.text or '' for t in el.findall('.//w:t',NS))
def save(name,obj):(OUT/name).write_text(json.dumps(obj,ensure_ascii=False,indent=2,default=str),encoding='utf-8')
def sha(p):return hashlib.sha256(p.read_bytes()).hexdigest()
registry=read_registry(APP/'DEV'/'ВЗМК_DEV.xlsx')
projects={}
for name in registry.projects:
    try:projects[name]=vars(select_project(registry,name))
    except Exception as e:projects[name]={'error':str(e)}
save('dev_project_contexts.json',projects)
inventory=json.loads((APP/'validation'/'stage5_uzk_audit_20261007_stop'/'file_inventory.json').read_text(encoding='utf-8'))
candidates=[]
for item in inventory:
    if not item['classification'].startswith('first-passport'):continue
    record={k:item[k] for k in ('path','sha256','size_bytes','modified_local_utc_plus_04')}
    text=item.get('content_excerpt','')
    match=re.search(r'Акт\s*№\s*(.*?)\s+от\s+(\d{2}\.\d{2}\.\d{4})',text)
    record['observed_act']=match.group(1) if match else None
    record['observed_date']=match.group(2) if match else None
    object_matches=[n for n,p in projects.items() if p.get('object_name') and p['object_name'] in text]
    record['DEV_object_text_matches']=object_matches
    record['number_matches']=[n for n in object_matches if projects[n].get('first_number')==record['observed_act']]
    record['classification']='first-passport candidate, other project; not selected'
    if Path(item['path'])==ROOT/'Гараж казахстан'/'Паспорт'/'1_2026 - УЗК.docx':
        assert record['observed_act']==projects['Гараж Казахстан']['first_number']
        record['classification']='selected project first-passport template: Гараж Казахстан'
    record['hash_verification']='prior audit snapshot; live recheck only if structurally inspected'
    candidates.append(record)
save('first_passport_classification.json',candidates)
paths=[ROOT/'Гараж казахстан'/'Паспорт'/'1_2026 - УЗК.docx',
       ROOT/'Гараж казахстан'/'Паспорт'/'2 от 15.09.2026'/'2_2026 - УЗК.docx',
       ROOT/'ТЭМПО'/'2 выдача'/'Паспорта'/'1 от 05.08.2026'/'1_2_2026 - УЗК.docx']
results=[]
for path in paths:
    original=next(i for i in inventory if i['path']==str(path))
    before=sha(path)
    assert before==original['sha256'],'STOP DOCX changed since prior inventory '+str(path)
    with ZipFile(path) as z:
        document=ET.fromstring(z.read('word/document.xml'))
        body=document.find('w:body',NS)
        result={'path':str(path),'sha256_before':before,'paragraphs':[], 'tables':[],
                'sections':[ET.tostring(x,encoding='unicode') for x in document.findall('.//w:sectPr',NS)],
                'total_paragraph_elements':len(document.findall('.//w:p',NS)),
                'content_controls':[{'tag':(x.find('w:sdtPr/w:tag',NS).get(W+'val') if x.find('w:sdtPr/w:tag',NS) is not None else None),'text':txt(x)} for x in document.findall('.//w:sdt',NS)],
                'textboxes':[txt(x) for x in document.findall('.//w:txbxContent',NS)],
                'drawing_count':len(document.findall('.//w:drawing',NS)),
                'pict_count':len(document.findall('.//w:pict',NS)),
                'headers_footers':{}, 'media':[]}
        for idx,p in enumerate(body.findall('w:p',NS),1):result['paragraphs'].append({'paragraph':idx,'text':txt(p),'xml':ET.tostring(p,encoding='unicode')})
        for ti,t in enumerate(body.findall('w:tbl',NS),1):
            table={'table':ti,'grid_columns':[x.get(W+'w') for x in t.findall('w:tblGrid/w:gridCol',NS)],
                   'properties':ET.tostring(t.find('w:tblPr',NS),encoding='unicode'),'rows':[]}
            for ri,r in enumerate(t.findall('w:tr',NS),1):
                row={'row':ri,'repeat_header':r.find('w:trPr/w:tblHeader',NS) is not None,'properties':ET.tostring(r.find('w:trPr',NS),encoding='unicode') if r.find('w:trPr',NS) is not None else None,'cells':[]}
                grid=1
                for ci,c in enumerate(r.findall('w:tc',NS),1):
                    span=c.find('w:tcPr/w:gridSpan',NS);vm=c.find('w:tcPr/w:vMerge',NS)
                    gs=int(span.get(W+'val')) if span is not None else 1
                    row['cells'].append({'physical_cell':ci,'grid_start':grid,'grid_span':gs,'vmerge':vm.get(W+'val','continue') if vm is not None else None,
                                         'paragraph_texts':[txt(p) for p in c.findall('w:p',NS)],'text':txt(c),'xml':ET.tostring(c,encoding='unicode')})
                    grid+=gs
                table['rows'].append(row)
            result['tables'].append(table)
        for n in z.namelist():
            if re.match(r'word/(header|footer)\d+\.xml$',n):
                xml=ET.fromstring(z.read(n));result['headers_footers'][n]={'text':txt(xml),'field_codes':[x.text for x in xml.findall('.//w:instrText',NS)],'xml':ET.tostring(xml,encoding='unicode')}
            if n.startswith('word/media/'):
                data=z.read(n);result['media'].append({'part':n,'bytes':len(data),'sha256':hashlib.sha256(data).hexdigest()})
        result['document_fields']=[x.text for x in document.findall('.//w:instrText',NS)]
        result['settings_xml']=z.read('word/settings.xml').decode('utf-8')
    result['sha256_after']=sha(path);assert before==result['sha256_after']
    results.append(result)
save('word_structure.json',results)
save('word_inputs.json',[{'path':str(p),'sha256':sha(p)} for p in paths])
print(json.dumps([{'path':x['path'],'paragraphs':[p['text'] for p in x['paragraphs']],
                   'tables':[{'table':t['table'],'grid_columns':len(t['grid_columns']),'rows':len(t['rows']),
                              'first_rows':[[c['text'] for c in r['cells']] for r in t['rows'][:4]]} for t in x['tables']],
                   'media':x['media'],'fields':x['document_fields'],'headers_footers':{k:v['text'] for k,v in x['headers_footers'].items()}} for x in results],ensure_ascii=True))
