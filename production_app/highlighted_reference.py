"""Format-preserving copies of the user's marked FIRST VIK/UZK DOCX.

VIK: replace ONLY the article names between the two uniquely identified
anchors in act 2 and two pairs of existing numbered/date content controls.
UZK: keep the exact old header, equipment, personnel, normative prose,
column widths, styles and signature layout; replace only the existing
data rows' mark, L/S, control date, frequency and angle, plus act number/date.

An inherited old statement such as "годен", "ДНО" or a name is copied
VERBATIM for visual/layout preview. It is *NOT* an actual finding for the
new shipment. This module must not be used to bypass current-inspection
verification during the official release or to label a preview "approved".

This only writes to an explicitly different NEW target DOCX. Source files
are pinned by SHA before/after, and original first passports never change.
"""
from collections import OrderedDict
from copy import deepcopy
from datetime import date,datetime
from pathlib import Path
import re
import zipfile
from decimal import Decimal

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from .storage import Stop,sha

VIK_LEFT='НШ, ТШ, СШ соединения '
VIK_RIGHT=' МП в соответствии'
UZK_DESIGN_COLS=(0,1,4,5,6)
SAFE_NUMBER=re.compile(r'^[1-9][0-9]*/20[0-9]{2}$')
WORD_DATE=re.compile(r'^[0-3][0-9]\.[01][0-9]\.20[0-9]{2}$')


def _inputs(source,target,number,control_date):
    original=Path(source).resolve(strict=True)
    output=Path(target).resolve(strict=False)
    if (original.suffix.lower()!='.docx' or not original.is_file()
            or output.suffix.lower()!='.docx' or output==original
            or output.exists() or not output.parent.is_dir()):
        raise Stop('DOCX первого паспорта копируется только в НОВЫЙ '
                   'отдельный файл; исходники менять нельзя.')
    if not isinstance(number,str) or not SAFE_NUMBER.fullmatch(number):
        raise Stop('Номер акта должен быть взят из реестра ВЗМК (N/YYYY).')
    if isinstance(control_date,datetime):control_date=control_date.date()
    if isinstance(control_date,date):control_date=control_date.isoformat()
    try: day=date.fromisoformat(control_date)
    except (ValueError,TypeError) as exc:
        raise Stop('Дата акта отсутствует в проверенных данных реестра.') from exc
    return original,output,number,day.strftime('%d.%m.%Y')


def _text_nodes(element):
    return list(element.iter(qn('w:t')))


def _plain(element):
    return ''.join(n.text or '' for n in _text_nodes(element))


def _clear_yellow(run):
    properties=run.find(qn('w:rPr'))
    if properties is not None:
        for node in list(properties.findall(qn('w:highlight'))):
            if node.get(qn('w:val'))=='yellow':properties.remove(node)


def _save_preserving_parts(document,source,output,*,relationships=False):
    """Keep opaque source parts byte-identical; edit only explicit slots."""
    payload={ 'word/document.xml':document.element.xml.encode('utf-8') }
    if relationships:payload['word/_rels/document.xml.rels']=document.part.rels.xml
    with zipfile.ZipFile(source) as original,zipfile.ZipFile(output,'w') as result:
        for info in original.infolist():
            result.writestr(info,payload.get(info.filename,original.read(info.filename)))


def _set_sdt(document,tag,value,expected):
    matches=[]
    for node in document.element.body.iter(qn('w:sdt')):
        properties=node.find(qn('w:sdtPr'))
        if properties is None:continue
        tag_el=properties.find(qn('w:tag'))
        if tag_el is None or tag_el.get(qn('w:val'))!=tag:continue
        body=node.find(qn('w:sdtContent'))
        if body is None or not _text_nodes(body):
            raise Stop('Word: повреждено существующее поле '+tag)
        matches.append(body)
    if len(matches)!=expected:
        raise Stop('Word: ожидается '+str(expected)+' полей '+tag+
                   ', найдено '+str(len(matches))+'; форма отличается.')
    for body in matches:
        nodes=_text_nodes(body)
        suffix='-1' if tag=='PASSPORT_NO' and re.fullmatch(r'[1-9][0-9]*/20[0-9]{2}-1',_plain(body).strip()) else ''
        nodes[0].text=str(value)+suffix
        for n in nodes[1:]:n.text=''
        for run in body.iter(qn('w:r')):_clear_yellow(run)
    return len(matches)


def _overwrite_slice(element,left,right,replacement):
    nodes=_text_nodes(element)
    text=''.join(n.text or '' for n in nodes)
    p=text.find(left)
    if p<0 or text.find(left,p+len(left))>=0:
        raise Stop('ВИК: не найден единственный левый якорь.')
    start=p+len(left)
    end=text.find(right,start)
    if end<0:
        raise Stop('ВИК: не найден правый якорь переменной части.')
    if start>=end:
        raise Stop('ВИК: пустое исходное наименование.')
    cursor=0
    inserted=False
    touched=set()
    for t in nodes:
        old=t.text or ''
        lo,hi=cursor,cursor+len(old)
        cursor=hi
        if hi<=start or lo>=end:continue
        a=max(0,start-lo)
        b=min(len(old),end-lo)
        if b<=a:continue
        head=old[:a]
        tail=old[b:]
        t.text=head+(replacement if not inserted else '')+tail
        inserted=True
        touched.add(t.getparent())
    if not inserted:
        raise Stop('ВИК: невозможно заменить исходный текст.')
    # Yellow highlight was *editing guidance*, not part of the issued copy.
    for run in touched:_clear_yellow(run)


def _replace_vik_names(document,names):
    unique=[]
    for raw in names:
        value=str(raw or '').strip()
        if not value or '{{' in value or '}}' in value:
            raise Stop('В ВОМ обнаружено пустое/неверное наименование изделия.')
        if value not in unique:unique.append(value)
    if not unique:raise Stop('В ВОМ нет наименований изделий для ВИК.')
    matched=[]
    for para in document.element.body.iter(qn('w:p')):
        value=_plain(para)
        if VIK_LEFT in value and VIK_RIGHT in value:
            matched.append(para)
    if len(matched)!=1:
        raise Stop('ВИК: в первом документе ожидается ровно одно место '
                   'для наименований второго акта; найдено '+str(len(matched)))
    _overwrite_slice(matched[0],VIK_LEFT,VIK_RIGHT,', '.join(unique))
    return unique


def _unset_old_mark_link(cell):
    """Do not carry a K1-5.pdf hyperlink into a NEW mark Бп2-1."""
    for node in list(cell._tc.iter(qn('w:hyperlink'))):
        parent=node.getparent()
        pos=parent.index(node)
        for child in list(node):
            parent.insert(pos,child)
            pos+=1
        parent.remove(node)


def _set_cell_keep_layout(cell,value):
    """Replace cell text without destroying original borders/merges/fonts."""
    runs=list(cell._tc.iter(qn('w:r')))
    if not runs:
        raise Stop('УЗК: строка первого документа не содержит текста.')
    first=runs[0]
    # Word can inherit the marked editing colour from paragraph run properties.
    # Clear it only in this replaced cell, retaining every other cell's style.
    for paragraph in cell._tc.iter(qn('w:p')):
        properties=paragraph.find(qn('w:pPr'))
        if properties is not None:_clear_yellow(properties)
    for run in runs:
        for child in list(run):
            if child.tag in (qn('w:t'),qn('w:br'),qn('w:cr'),qn('w:tab')):
                run.remove(child)
        _clear_yellow(run)
    pieces=str(value).split('\n')
    for i,piece in enumerate(pieces):
        if i:first.append(OxmlElement('w:br'))
        t=OxmlElement('w:t')
        t.text=piece
        if piece[:1].isspace() or piece[-1:].isspace():
            t.set(qn('xml:space'),'preserve')
        first.append(t)


def _replace_uzk_heading(document,number,control_date):
    matches=[]
    for paragraph in document.paragraphs:
        text=paragraph.text
        if text.startswith('Акт № ') and ' от ' in text:
            matches.append(paragraph)
    if len(matches)!=1:
        raise Stop('УЗК: не найден единственный заголовок «Акт № ... от».')
    p=matches[0]
    original=p.text
    match=re.fullmatch(r'Акт № ([0-9]+/20[0-9]{2}) от ([0-9]{2}\.[0-9]{2}\.20[0-9]{2})',
                       original)
    if not match:
        raise Stop('УЗК: не распознан исторический номер или дата акта.')
    _overwrite_slice(p._p,'Акт № ',' от ',number)
    full=_plain(p._p)
    start=full.find(' от ')+len(' от ')
    if start<4:
        raise Stop('УЗК: ошибка исходной даты в заголовке.')
    nodes=_text_nodes(p._p)
    cursor=0
    changed=False
    for node in nodes:
        text=node.text or ''
        stop=cursor+len(text)
        if stop>start:
            index=start-cursor if cursor<start else 0
            node.text=text[:index]+(control_date if not changed else '')
            changed=True
            _clear_yellow(node.getparent())
        cursor=stop
    if not changed:
        raise Stop('УЗК: не заменена дата акта.')


def _normalize_design(design_rows):
    if not isinstance(design_rows,(list,tuple)) or not design_rows:
        raise Stop('УЗК: нет подтверждённых плановых марок/характеристик.')
    grouped=OrderedDict()
    for record in design_rows:
        if not isinstance(record,dict):
            raise Stop('УЗК: некорректная строка расчёта.')
        mark=str(record.get('mark') or '').strip()
        if not mark or '{{' in mark or '}}' in mark:
            raise Stop('УЗК: пустая/опасная марка.')
        data=[]
        for field in ('length_mm','thickness_mm','frequency_mhz','angle_deg'):
            text=str(record.get(field) or '').strip()
            if not text or not re.fullmatch(r'[0-9]+(?:[.,][0-9]+)?',text):
                raise Stop('УЗК: нет расчётного значения '+field+' для '+mark)
            n=Decimal(text.replace(',','.'))
            maximum={'frequency_mhz':Decimal('20'),'angle_deg':Decimal('90')}.get(field)
            if n<=0 or maximum is not None and n>maximum:
                raise Stop('УЗК: недопустимое '+field+' для '+mark)
            data.append(text.replace('.',','))
        group=grouped.setdefault(mark,[])
        group.append(tuple(data))
        if len(group)>10:
            raise Stop('УЗК: слишком много отдельных толщин по марке '+mark)
    if len(grouped)>5000:
        raise Stop('УЗК: превышен допустимый размер таблицы.')
    return grouped


def _strip_unused_old_links(document):
    used=set()
    for e in document.element.body.iter():
        for name,val in e.attrib.items():
            if name in (qn('r:id'),qn('r:embed'),qn('r:link')):
                used.add(val)
    for rid,rel in list(document.part.rels.items()):
        if (rel.reltype.endswith('/hyperlink') and rid not in used):
            del document.part.rels[rid]


def _copy_uzk_rows(document,design_rows,control_date):
    table=document.tables[0]
    if len(document.tables)<2 or len(table.rows)<6:
        raise Stop('УЗК: отличается структура таблицы первого паспорта.')
    header=' '.join(x.text for x in table.rows[4].cells)
    if 'Характеристика сварного соединения' not in header:
        raise Stop('УЗК: не найдена исходная шапка характеристик.')
    original=table.rows[5]
    if len(original.cells)!=9 or not original.cells[0].text.strip():
        raise Stop('УЗК: в первом документе неверная структура строк.')
    if original.cells[1]._tc is not original.cells[2]._tc:
        raise Stop('УЗК: характеристика должна занимать два объединённых столбца.')
    for old in list(table.rows[5:]):
        table._tbl.remove(old._tr)
    template=deepcopy(original._tr)
    # A complete mark row (including all welders) must stay on one page.
    # The original five-row heading is a contiguous group in this exact
    # user form; repeat it together when the dynamic body continues.
    from docx.oxml import OxmlElement
    for heading in table.rows[:5]:
        properties=heading._tr.get_or_add_trPr()
        if properties.find(qn('w:tblHeader')) is None:
            properties.append(OxmlElement('w:tblHeader'))
    properties=template.get_or_add_trPr()
    if properties.find(qn('w:cantSplit')) is None:
        properties.append(OxmlElement('w:cantSplit'))
    mapped=_normalize_design(design_rows)
    for mark,records in mapped.items():
        table._tbl.append(deepcopy(template))
        row=table.rows[-1]
        if len(row.cells)!=9:
            raise Stop('УЗК: при добавлении марки нарушена структура таблицы.')
        _unset_old_mark_link(row.cells[0])
        _set_cell_keep_layout(row.cells[0],mark)
        _set_cell_keep_layout(row.cells[1],'\n'.join(
            'L='+length+'х'+thick+' мм' for length,thick,_,_ in records))
        _set_cell_keep_layout(row.cells[4],control_date)
        _set_cell_keep_layout(row.cells[5],'\n'.join(f for _,_,f,_ in records))
        _set_cell_keep_layout(row.cells[6],'\n'.join(a for _,_,_,a in records))
    _strip_unused_old_links(document)
    if len(table.rows)!=5+len(mapped):
        raise Stop('УЗК: после подстановки марок неверное число строк.')
    return tuple(mapped)



def supports_vik_first(first):
    """Recognize exactly this first-passport two-act VIK layout; read-only."""
    try:
        d=Document(first)
        expected={'PASSPORT_NO':0,'VIK_DATE':0}
        number_values=[]
        for node in d.element.body.iter(qn('w:sdt')):
            pr=node.find(qn('w:sdtPr'))
            if pr is None:continue
            tag=pr.find(qn('w:tag'))
            if tag is None:continue
            key=tag.get(qn('w:val'))
            if key in expected:expected[key]+=1
            if key=='PASSPORT_NO':number_values.append(_plain(node.find(qn('w:sdtContent'))).strip())
        matches=sum(1 for p in d.element.body.iter(qn('w:p'))
                    if VIK_LEFT in _plain(p) and VIK_RIGHT in _plain(p))
        return (expected=={'PASSPORT_NO':2,'VIK_DATE':2} and matches==1
                and all(re.fullmatch(r'[1-9][0-9]*/20[0-9]{2}(?:-1)?',n) for n in number_values))
    except Exception:
        return False


def supports_uzk_first(first):
    """Recognize exact table/heading pattern; never an arbitrary DOCX."""
    try:
        d=Document(first)
        if len(d.tables)<2 or len(d.tables[0].rows)<6:
            return False
        t=d.tables[0]
        return (
            len(t.rows[5].cells)==9
            and t.rows[5].cells[1]._tc is t.rows[5].cells[2]._tc
            and 'Характеристика сварного соединения' in
            ' '.join(c.text for c in t.rows[4].cells)
            and sum(1 for p in d.paragraphs
                    if re.fullmatch(
                        r'Акт № [0-9]+/20[0-9]{2} от '
                        r'[0-9]{2}\.[0-9]{2}\.20[0-9]{2}',p.text))==1
        )
    except Exception:
        return False


def create_vik_copy(first,target,*,number,vik_date,item_names):
    source,output,number,day=_inputs(first,target,number,vik_date)
    old_sha=sha(source)
    d=Document(source)
    _set_sdt(d,'PASSPORT_NO',number,2)
    _set_sdt(d,'VIK_DATE',day,2)
    names=_replace_vik_names(d,item_names)
    _save_preserving_parts(d,source,output)
    if sha(source)!=old_sha:
        raise Stop('ВИК: исходный первый паспорт изменился.')
    return {'file':str(output),'source_sha256':old_sha,
            'item_names':names,'number':number,'date':day,
            'status':'PREVIEW_FROM_HISTORICAL_FORM_UNVERIFIED_CURRENT_CONTROL'}


def create_uzk_copy(first,target,*,number,uzk_date,design_rows):
    source,output,number,day=_inputs(first,target,number,uzk_date)
    old_sha=sha(source)
    d=Document(source)
    _replace_uzk_heading(d,number,day)
    marks=_copy_uzk_rows(d,design_rows,day)
    _save_preserving_parts(d,source,output,relationships=True)
    if sha(source)!=old_sha:
        raise Stop('УЗК: исходный первый паспорт изменился.')
    return {'file':str(output),'source_sha256':old_sha,
            'marks':marks,'number':number,'date':day,
            'status':'PREVIEW_FROM_HISTORICAL_FORM_UNVERIFIED_CURRENT_CONTROL'}
