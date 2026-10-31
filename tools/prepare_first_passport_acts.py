"""Prepare new cleared VIK/UZK forms from explicitly SHA-pinned references.

This reviewed two-act/eight-column layout adapter is deliberately fail closed.
No historical outcome, personnel identity or signature is approved by copying.
Only word/document.xml changes; opaque package parts remain byte-identical.
"""
import argparse
from copy import deepcopy
import hashlib
from pathlib import Path
import re
import sys
from lxml import etree as ET

sys.path.insert(0,str(Path(__file__).resolve().parents[1]))
from production_app.native_acts import W,text,set_text,patch_document
from production_app.storage import Stop


def replace_literal(node,old,new):
    """Replace even split runs, retaining all untouched text/run styles."""
    for paragraph in node.iter(W+'p'):
        nodes=list(paragraph.iter(W+'t'));whole=''.join(n.text or '' for n in nodes)
        start=whole.find(old)
        while start>=0:
            end=start+len(old);offset=0;first=True
            for item in nodes:
                value=item.text or '';a=offset;b=offset+len(value);offset=b
                if b<=start or a>=end:continue
                left=value[:max(0,start-a)];right=value[max(0,end-a):]
                item.text=left+(new if first else '')+right;first=False
            whole=''.join(n.text or '' for n in nodes);start=whole.find(old)


def _cell(table,row,col=0):return table.findall(W+'tr')[row].findall(W+'tc')[col]


def _check(root):
    if list(root.iter(W+'drawing')) or list(root.iter(W+'pict')):
        raise Stop('В исходнике есть рисунки/подписи: нужна отдельная проверка.')
    if list(root.iter(W+'ins')) or list(root.iter(W+'del')):
        raise Stop('В исходнике есть незавершённые исправления.')


def adapt_vik(root):
    _check(root);body=root.find(W+'body');items=list(body)
    if len(items)!=25 or [i for i,x in enumerate(items) if x.tag==W+'tbl']!=[0,3,5,7,12,13,18,22]:
        raise Stop('Неизвестная структура двух актов ВИК. Замены не выполнялись.')
    if 'Визуального и измерительного контроля' not in text(items[0]):
        raise Stop('Не найден заголовок исходного ВИК.')
    replace_literal(root,'1/2026-1','{{VIK_PROCESS_NUMBER}}')
    replace_literal(root,'1/2026','{{NUMBER}}')
    for control in root.iter(W+'sdt'):
        if text(control).strip()=='{{VIK_PROCESS_NUMBER}}':
            tag=control.find(W+'sdtPr/'+W+'tag')
            if tag is not None:tag.set(W+'val','VIK_PROCESS_NUMBER')
    replace_literal(root,'02.09.2026','{{VIK_DATE}}')
    set_text(items[2],'1. В соответствии с договором {{CONTRACT}} на поставку продукции')
    set_text(_cell(items[3],4),'{{VIK_WELDERS}}')
    for row in (5,6):set_text(_cell(items[3],row),'')
    set_text(_cell(items[3],7),'Металлоконструкции по шифру: {{CIPHER}}')
    set_text(_cell(items[3],8),'Приложение №1. Марки: {{VIK_MARKS}}')
    set_text(_cell(items[7],0),'{{VIK_DEFECTS}}')
    set_text(_cell(items[7],2),'Конструкции изготовлены из проката и листовой стали {{STEEL}}')
    set_text(items[10],'{{VIK_CONCLUSION}}')
    set_text(_cell(items[12],0),'Контроль выполнил:\n{{VIK_INSPECTOR}}')
    set_text(_cell(items[12],1),'Руководитель работ по визуальному и измерительному контролю:\n{{VIK_HEAD}}')
    set_text(items[15],'{{OBJECT}} ({{CIPHER}})')
    set_text(_cell(items[18],0),'{{VIK_WELDERS}}')
    for row in (1,2):set_text(_cell(items[18],row),'')
    set_text(_cell(items[18],4),'НШ, ТШ, СШ соединения {{CONNECTIONS}}. Марки: {{VIK_MARKS}}. МП в соответствии с требованиями технологии сварки')
    set_text(_cell(items[22],0),'{{VIK_PROCESS_CONCLUSION}}')
    set_text(_cell(items[22],2),'Контроль выполнил:\n{{VIK_INSPECTOR}}')
    set_text(_cell(items[22],3),'Руководитель работ по визуальному и измерительному контролю:\n{{VIK_HEAD}}')
    set_text(items[23],'Основание: {{EVIDENCE}}\nОтветственный за комплект: {{RESPONSIBLE}}')
    # After clearing long historical personnel blocks, natural pagination
    # can orphan the second act's heading at the end of act1. Preserve two
    # complete independent acts with an explicit source-derived page break.
    separator=ET.Element(W+'p');props=ET.SubElement(separator,W+'pPr')
    ET.SubElement(props,W+'pageBreakBefore');ET.SubElement(props,W+'keepNext')
    spacing=ET.SubElement(props,W+'spacing');spacing.set(W+'before','0');spacing.set(W+'after','0');spacing.set(W+'line','20');spacing.set(W+'lineRule','exact')
    body.insert(list(body).index(items[13]),separator)


def adapt_uzk(root):
    _check(root);body=root.find(W+'body');items=list(body)
    if len(items)!=11 or [i for i,x in enumerate(items) if x.tag==W+'tbl']!=[6,8]:
        raise Stop('Неизвестная структура первого УЗК. Замены не выполнялись.')
    table=items[6];rows=table.findall(W+'tr')
    if len(rows)!=16 or len(rows[4].findall(W+'tc'))!=8 or len(rows[5].findall(W+'tc'))!=8:
        raise Stop('Изменилась восьмиколоночная таблица первого УЗК.')
    replace_literal(items[4],'1/2026','{{NUMBER}}')
    replace_literal(items[4],'04.09.2026','{{UZK_DATE}}')
    set_text(_cell(table,1,1),'{{OBJECT}} ({{CIPHER}})')
    # Technical equipment/calibration requisites remain source-derived. A
    # human still approves current applicability of this prepared form.
    details=deepcopy(table)
    for row in list(details.findall(W+'tr')):details.remove(row)
    header=deepcopy(rows[4]);prototype=deepcopy(rows[5])
    props=header.find(W+'trPr')
    if props is None:props=ET.Element(W+'trPr');header.insert(0,props)
    ET.SubElement(props,W+'tblHeader');ET.SubElement(props,W+'cantSplit')
    for cell in prototype.findall(W+'tc'):set_text(cell,'')
    set_text(prototype.findall(W+'tc')[0],'{{UZK_NATIVE_ROW}}')
    set_text(prototype.findall(W+'tc')[1],'{{UZK_DESIGN}}')
    details.extend((header,prototype))
    for row in rows[4:]:table.remove(row)
    # Separate the source requisites and data tables. Only the column header
    # repeats; no first-page equipment/title block consumes later pages.
    separator=ET.Element(W+'p');ppr=ET.SubElement(separator,W+'pPr')
    spacing=ET.SubElement(ppr,W+'spacing');spacing.set(W+'before','0');spacing.set(W+'after','0');spacing.set(W+'line','20');spacing.set(W+'lineRule','exact')
    position=list(body).index(table);body.insert(position+1,separator);body.insert(position+2,details)
    set_text(_cell(items[8],0,2),'{{UZK_INSPECTOR}}')
    set_text(_cell(items[8],1),'Квалификация и удостоверение — в подтверждённых реквизитах исполнителя')
    set_text(items[9],'Основание: {{EVIDENCE}}\nОтветственный за комплект: {{RESPONSIBLE}}')


def prepare(source,target,role,expected_sha):
    source=Path(source)
    actual=hashlib.sha256(source.read_bytes()).hexdigest()
    if actual!=expected_sha:raise Stop('Изменился проверенный исходный '+role)
    patch_document(source,target,adapt_vik if role=='vik' else adapt_uzk)
    if hashlib.sha256(source.read_bytes()).hexdigest()!=actual:raise Stop('Оригинал изменился во время подготовки формы.')


def main():
    p=argparse.ArgumentParser(description=__doc__)
    for key in ('vik','uzk','vik-sha256','uzk-sha256','output'):p.add_argument('--'+key,required=True)
    args=p.parse_args();out=Path(args.output);out.mkdir(exist_ok=False,parents=True)
    prepare(args.vik,out/'ВИК.docx','vik',args.vik_sha256)
    prepare(args.uzk,out/'УЗК.docx','uzk',args.uzk_sha256)


if __name__=='__main__':main()
