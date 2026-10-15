import json,zipfile,hashlib,re
from pathlib import Path
from xml.etree import ElementTree as E
OUT=Path(__file__).resolve().parent;ROOT=OUT.parents[1];P=ROOT.parent/'Гараж казахстан/Паспорт';W='{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
def sha(p):return hashlib.sha256(p.read_bytes()).hexdigest()
paths=[P/'1_2026 - ВИК.docx',P/'9 от 06.10.2026/9_2026 - ВИК.docx',P/'1_2026 - УЗК.docx',P/'9 от 06.10.2026/9_2026 - УЗК.docx',ROOT/'e2e_user_stage2/validation/user_e2e/Документы/Гараж казахстан/Паспорт/9-2026 от 06.10.2026/9_2026 - ВИК.docx']
result=[]
for p in paths:
    with zipfile.ZipFile(p) as z:
        x=E.fromstring(z.read('word/document.xml'));paras=[''.join(t.text or '' for t in a.iter(W+'t')) for a in x.iter(W+'p')];sdt=[]
        for a in x.iter(W+'sdt'):
            tag=a.find('.//'+W+'tag');sdt.append({'tag':tag.get(W+'val') if tag is not None else None,'text':''.join(t.text or '' for t in a.iter(W+'t'))})
        result.append({'path':str(p),'sha256':sha(p),'paragraphs':paras,'content_controls':sdt})
(OUT/'supporting_docx.json').write_text(json.dumps(result,ensure_ascii=False,indent=2),encoding='utf8')
for d in result:
    print(d['path']);print('\n'.join(t for t in d['paragraphs'] if re.search(r'АКТ|Акт|[19]/2026|\d\d\.\d\d\.2026|ТЕСТ',t)))
    print('controls',d['content_controls'])
