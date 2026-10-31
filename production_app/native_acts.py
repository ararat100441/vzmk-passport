"""Opt-in eight-column first-passport layout; geometry is not a test verdict.

Only a prepared row carrying UZK_NATIVE_ROW is expanded. All unrelated OOXML
parts and the reference cell/paragraph/run properties remain unchanged.
"""
from collections import defaultdict
from copy import deepcopy
from datetime import date
from decimal import Decimal
from pathlib import Path
import re
import zipfile
from lxml import etree as ET

from .storage import Stop

W='{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
ET.register_namespace('w',W[1:-1])
NATIVE_MARKER='UZK_NATIVE_ROW'


def text(node):return ''.join(t.text or '' for t in node.iter(W+'t'))


def set_text(node,value):
    """Change one paragraph/cell slot, retaining its source text style."""
    if node.tag==W+'tc':
        paragraphs=node.findall(W+'p')
        if not paragraphs:raise Stop('В форме УЗК отсутствует абзац ячейки.')
        paragraph=paragraphs[0]
        for other in paragraphs[1:]:node.remove(other)
    else:paragraph=node
    style=next((deepcopy(r.find(W+'rPr')) for r in paragraph.iter(W+'r')
                if r.find(W+'rPr') is not None),None)
    for child in list(paragraph):
        if child.tag!=W+'pPr':paragraph.remove(child)
    run=ET.SubElement(paragraph,W+'r')
    if style is not None:run.append(style)
    for n,line in enumerate(str(value).split('\n')):
        if n:ET.SubElement(run,W+'br')
        t=ET.SubElement(run,W+'t');t.set('{http://www.w3.org/XML/1998/namespace}space','preserve');t.text=line


def patch_document(source,target,transform):
    source,target=Path(source),Path(target)
    if source.resolve()==target.resolve() or target.exists():
        raise Stop('Нужен новый отдельный DOCX, перезапись запрещена.')
    with zipfile.ZipFile(source) as package:
        root=ET.fromstring(package.read('word/document.xml'))
        transform(root)
        payload=ET.tostring(root,encoding='utf-8',xml_declaration=True)
        with zipfile.ZipFile(target,'w') as result:
            for info in package.infolist():
                result.writestr(info,payload if info.filename=='word/document.xml' else package.read(info.filename))


def _parameters(value):
    try:return tuple(Decimal(s.strip().replace(',','.')) for s in re.split(r'[;|/]+',str(value)))
    except Exception as exc:raise Stop('Некорректные фактические параметры УЗК.') from exc


def native_rows(facts,*,draft=False):
    """One row per selected mark, with independent wall/flange lengths.

    Actual welders/defects/conclusions originate ONLY from supplied current
    detailed records. A design preview leaves them explicitly unconfirmed.
    """
    design=defaultdict(list)
    for row in facts.get('uzk_design',[]):design[row['mark']].append(row)
    observed=defaultdict(list)
    for row in facts.get('uzk_details',[]):observed[row['mark']].append(row)
    result=[]
    for mark,parts in design.items():
        quantities={r['quantity'] for r in parts};profiles={r['profile'] for r in parts}
        if len(quantities)!=1 or len(profiles)!=1:raise Stop('Неоднозначная геометрия УЗК: '+mark)
        quantity=next(iter(quantities));profile=next(iter(profiles))
        joint=f'{profile}; {quantity} шт.\n'+'\n'.join('L='+r['length_mm'].replace('.',',')+'х'+r['thickness_mm'].replace('.',',')+' мм' for r in parts)
        freq='\n'.join(r['frequency_mhz'].replace('.',',') for r in parts)
        angle='\n'.join(r['angle_deg'] for r in parts)
        if draft:
            welders=day=defects=conclusion='НЕТ ДАННЫХ'
        else:
            if len(observed[mark])!=1:
                raise Stop('Для формы первого паспорта нужен один подтверждённый набор текущих результатов УЗК по марке '+mark)
            current=observed[mark][0]
            for key in ('welders','date','defects','conclusion','basis'):
                if not isinstance(current.get(key),str) or not current[key].strip():
                    raise Stop('Нет текущего результата УЗК '+key+': '+mark)
            for key,expected in (('frequency_mhz',tuple(Decimal(r['frequency_mhz']) for r in parts)),
                                 ('angle_deg',tuple(Decimal(r['angle_deg']) for r in parts))):
                actual=_parameters(current.get(key,''))
                if len(actual)==1:actual=actual*len(expected)
                if actual!=expected:
                    raise Stop('Фактический параметр '+key+' не соответствует принятому правилу профиля: '+mark)
            welders=current['welders'];defects=current['defects'];conclusion=current['conclusion']
            try:day=date.fromisoformat(current['date']).strftime('%d.%m.%Y')
            except ValueError as exc:raise Stop('Неверная текущая дата УЗК: '+mark) from exc
        result.append((mark,joint,welders,day,freq,angle,defects,conclusion))
    # No controlled mark may silently disappear from the generated act.
    selected={str(m['mark']) for m in facts.get('marks',[]) if Decimal(str(m.get('uzk_quantity',0)))>0}
    if set(design)!=selected:raise Stop('Проектная геометрия УЗК не охватывает выбранные марки.')
    return result


def expand_native_uzk(source,target,facts,*,draft=False):
    rows=native_rows(facts,draft=draft)
    def transform(root):
        parents={child:parent for parent in root.iter() for child in parent}
        prototypes=[tr for tr in root.iter(W+'tr') if '{{UZK_NATIVE_ROW}}' in text(tr)]
        if len(prototypes)!=1:raise Stop('Нужна одна строка {{UZK_NATIVE_ROW}} в утверждённой форме УЗК.')
        prototype=prototypes[0];table=parents[prototype]
        if table.tag!=W+'tbl' or len(prototype.findall(W+'tc'))!=8:
            raise Stop('Форма первого УЗК должна иметь ровно восемь ячеек строки.')
        position=list(table).index(prototype)
        table.remove(prototype)
        for n,values in enumerate(rows):
            row=deepcopy(prototype)
            props=row.find(W+'trPr')
            if props is None:props=ET.Element(W+'trPr');row.insert(0,props)
            for height in list(props.findall(W+'trHeight')):props.remove(height)
            if props.find(W+'cantSplit') is None:ET.SubElement(props,W+'cantSplit')
            for cell,value in zip(row.findall(W+'tc'),values):set_text(cell,value)
            table.insert(position+n,row)
    patch_document(source,target,transform)
