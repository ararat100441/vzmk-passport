from pathlib import Path
import json
import sys
from zipfile import ZipFile
from lxml import etree as E

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT))
from uzk_stage5.inputs import sha, PROFILES, REGISTRY, UZK, CHECKPOINT, select_passport, project_sources

RUN = Path(__file__).resolve().parent
passport = select_passport('Гараж Казахстан', '9/2026')
project, catalog, template = project_sources(passport)
vom = project / 'Паспорт/9 от 06.10.2026/ВОМ отгрузочный счет 183  (5).xls'
existing = vom.parent / '9_2026 - УЗК.docx'
paths = [*ROOT.glob('*.py'), *ROOT.glob('*.cmd'), ROOT/'requirements.txt', ROOT/'settings.json',
         REGISTRY, UZK, CHECKPOINT, ROOT/'DEV/DEV_REGISTRY_CHECKPOINT.json',
         catalog, template, vom, existing, *PROFILES.parent.glob('*')]
for folder in ('passport_app', 'profiles_stage4', 'tests'):
    # Runtime node_modules are junctions, not accepted product sources.
    paths.extend(p for p in (ROOT/folder).glob('*') if p.is_file())
paths = [p for p in paths if p.is_file() and p.name not in ('stage5.py', 'Запустить Этап 5 DEV.cmd')]
protected = {}
for path in sorted(set(paths)):
    protected[str(path.resolve())] = sha(path)
(RUN/'protected_before.json').write_text(json.dumps(protected, ensure_ascii=False, indent=2), encoding='utf8')
(RUN/'product_before.json').write_text(json.dumps({'new_stage5_files': 'ABSENT_BEFORE_IMPLEMENTATION',
    'accepted_product': {k:v for k,v in protected.items() if any(x in Path(k).parts for x in ('passport_app','profiles_stage4'))
                         or Path(k).parent == ROOT}}, ensure_ascii=False, indent=2), encoding='utf8')
with ZipFile(template) as z:
    ns={'w':'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
    root=E.fromstring(z.read('word/document.xml'))
    import hashlib
    inventory={p.filename:hashlib.sha256(z.read(p.filename)).hexdigest() for p in z.infolist()}
    details={'parts':inventory,'section_xml':E.tostring(root.find('.//w:sectPr',ns),encoding='unicode'),
             'body': [{'index':i,'tag':E.QName(e).localname,'text':''.join(e.xpath('.//w:t/text()',namespaces=ns))}
                      for i,e in enumerate(root.find('w:body',ns))]}
    (RUN/'template_inventory.json').write_text(json.dumps(details,ensure_ascii=False,indent=2),encoding='utf8')
print(json.dumps({'run':str(RUN),'protected_files':len(protected),'passport':passport},ensure_ascii=False))
