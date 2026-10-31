"""Guarded edits of the accepted first template; all other ZIP parts unchanged."""
from copy import deepcopy
from io import BytesIO
from pathlib import Path
from zipfile import ZipFile
from lxml import etree as E
from .provenance import Stage7Error, sha
from .sources import TEMPLATE, PINS

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
NS = {'w': W}
BANNER = 'ТЕСТ / ЧЕРНОВИК. НЕ ДЛЯ ОФИЦИАЛЬНОЙ ВЫДАЧИ.'
STORAGE = 'Документы о качестве на металлопрокат, сварочные материалы и лакокрасочные материалы хранятся на предприятии'
SCOPE = ('ТЕСТ-комплект не подтверждает качество всей отгрузки. Результаты УЗК и АКЗ не подтверждены. '
         'АКЗ №0 охватывает только 3 назначенные марки / 3 изделия. Подпись и печать отсутствуют.')


def tag(name):
    return '{' + W + '}' + name


def text(node):
    return ''.join(node.xpath('.//w:t/text()', namespaces=NS))


def canonical(node):
    return E.tostring(node, method='c14n')


def parse(path):
    with ZipFile(path) as z:
        return E.fromstring(z.read('word/document.xml'))


def table_cells(root):
    return [[r.findall('w:tc', NS) for r in t.findall('w:tr', NS)] for t in root.findall('w:body/w:tbl', NS)]


def replace_exact(node, old, new):
    """Replace one verified local span across split runs, retaining run formatting."""
    full = text(node)
    if not old or full.count(old) != 1:
        raise Stage7Error('WORD_LOCAL_FIELD_AMBIGUOUS: ' + old)
    start, stop = full.index(old), full.index(old) + len(old)
    offset = 0
    first = True
    for t in node.xpath('.//w:t', namespaces=NS):
        raw = t.text or ''
        end = offset + len(raw)
        if end > start and offset < stop:
            left = raw[:max(0, start-offset)]
            right = raw[max(0, stop-offset):] if stop < end else ''
            t.text = left + (new if first else '') + right
            t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
            first = False
        offset = end
    if text(node) != full[:start] + new + full[stop:]:
        raise Stage7Error('WORD_LOCAL_REPLACEMENT_FAILED')


def set_paragraph(p, value, *, banner=False):
    visible = p.xpath('.//w:r[w:t]', namespaces=NS)
    rpr = visible[0].find('w:rPr', NS) if visible else p.find('w:pPr/w:rPr', NS)
    rpr = deepcopy(rpr) if rpr is not None else E.Element(tag('rPr'))
    for c in list(p):
        if c.tag != tag('pPr'):
            p.remove(c)
    if banner:
        for c in list(rpr):
            if c.tag in (tag('u'), tag('b'), tag('bCs')):
                rpr.remove(c)
        E.SubElement(rpr, tag('b'))
        E.SubElement(rpr, tag('u')).set(tag('val'), 'none')
    run = E.SubElement(p, tag('r'))
    run.append(rpr)
    t = E.SubElement(run, tag('t'))
    t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    t.text = value


def list_cell(cell, values):
    paragraphs = cell.findall('w:p', NS)
    if not paragraphs:
        raise Stage7Error('WORD_LIST_CELL_MISSING')
    prototype = deepcopy(paragraphs[0])
    for p in paragraphs:
        cell.remove(p)
    for v in values:
        p = deepcopy(prototype)
        set_paragraph(p, v)
        cell.append(p)


def references(manifest, usage):
    # Presentation is separate from status-bearing machine metadata. The AKZ
    # appendix stays mandatory in the verified pair, but is not listed here.
    labels = {'VIK': 'Акт ВИК', 'UZK': 'Акт УЗК',
              'AKZ': 'Сертификат качества антикоррозийной обработки'}
    result = []
    for record in manifest['records']:
        if usage not in record['usage'] or record['type'] not in labels:
            continue
        if record['official_confirmed'] is not False or record['status'] not in ('TEST', 'SAMPLE_ONLY'):
            raise Stage7Error('REFERENCE_STATUS_INVALID')
        result.append(f"{labels[record['type']]} № {record['number']} от {record['date']} г.")
    return result


def acts_paragraph(manifest, usage):
    return ', '.join(references(manifest, usage))


def contract_value(value):
    # The fixed adjacent label already contains №. Remove only that prefix
    # from the value, never from the contract's actual number or the label.
    value = value.strip()
    if value.startswith('№'):
        value = value[1:].lstrip()
    if not value or value.startswith('№') or any(c in value for c in '\r\n\t'):
        raise Stage7Error('CONTRACT_VALUE_INVALID')
    return value


EDITABLE = {(0,1,3),(0,9,0),(0,11,3),(0,12,1),(0,14,0),(0,15,1),
    (0,16,1),(0,17,1),(0,18,1),(0,19,0),(0,21,0),(0,23,1),(0,50,0),(0,53,0),(1,4,0)}


def structural_check(before, after):
    a, b = table_cells(before), table_cells(after)
    if len(a) != 2 or len(b) != 2 or [len(t) for t in a] != [56,7] or [len(t) for t in b] != [56,7]:
        raise Stage7Error('WORD_STRUCTURE_CHANGED')
    for ti, table in enumerate(a):
        for ri, row in enumerate(table):
            if len(row) != len(b[ti][ri]):
                raise Stage7Error('WORD_CELL_TOPOLOGY_CHANGED')
            for ci, cell in enumerate(row):
                out = b[ti][ri][ci]
                if (ti,ri,ci) not in EDITABLE and canonical(cell) != canonical(out):
                    raise Stage7Error('WORD_CONSTANT_CELL_CHANGED')
                if canonical(cell.find('w:tcPr',NS)) != canonical(out.find('w:tcPr',NS)):
                    raise Stage7Error('WORD_CELL_FORMAT_CHANGED')
    protected = [canonical(r) for r in before.findall('w:body/w:tbl',NS)[0].findall('w:tr',NS)[40:49]]
    protected_after = [canonical(r) for r in after.findall('w:body/w:tbl',NS)[0].findall('w:tr',NS)[40:49]]
    if protected != protected_after:
        raise Stage7Error('POINT12_CHANGED')
    for query, label in [('.//w:numPr','NUMBERING'),('.//w:sectPr','SECTION'),
                          ('.//w:instrText','MERGEFIELD'),('.//w:fldChar','FIELD'),('.//w:tblPr','TABLE_FORMAT'),('.//w:trPr','ROW_FORMAT')]:
        if [canonical(x) for x in before.xpath(query,namespaces=NS)] != [canonical(x) for x in after.xpath(query,namespaces=NS)]:
            raise Stage7Error('WORD_' + label + '_CHANGED')
    return {'status':'PASS','point12_xml_equal':True,'constant_cells_equal':True,
            'numbering_equal':True,'mergefield_equal':True,'tables_and_sections_equal':True}


def write_quality(manifest, output):
    if manifest['mode'] != 'DEV_TEST_ONLY' or manifest['official_issue'] is not False:
        raise Stage7Error('OFFICIAL_FORBIDDEN')
    if sha(TEMPLATE) != PINS[TEMPLATE]:
        raise Stage7Error('QUALITY_TEMPLATE_SHA_MISMATCH')
    root = parse(TEMPLATE)
    before = deepcopy(root)
    cells = table_cells(root)
    def cell(row,col=1,table=0):
        return cells[table][row-1][col-1]
    guards = {(12,1):'Заказ №',(12,3):'Договор №',(13,1):'Заказчик',
        (16,1):'Наименование конструкции:',(17,1):'Масса конструкций по чертежам изготовителя, кг:',
        (50,1):'Документ о качестве составлен на основании приемочных актов:',
        (53,1):'Согласно условиям договора на поставку и требованиям настоящего стандарта к документу о качестве прилагаются:'}
    for (row,col), expected in guards.items():
        if text(cell(row,col)) != expected:
            raise Stage7Error('QUALITY_TEMPLATE_LABEL_CHANGED')
    if text(cell(12,2)) != '1':
        raise Stage7Error('QUALITY_ORDER_CONSTANT_CHANGED')
    f = manifest['fields']
    replace_exact(cell(2,4),'1/2026',f['number'])
    replace_exact(cell(12,4),'№ TEST/26 от 01.01.2026',contract_value(f['contract']))
    replace_exact(cell(13,2),'ООО «СТК ВК»',f['customer'])
    replace_exact(cell(15), 'Архитектурно-строительные решения здания парковки подвижного состава',f['object'])
    replace_exact(cell(16,2),'SAMPLE-PROJECT-KMD',f['cipher'])
    mass = format(float(f['mass']),',.2f').replace(',',' ').replace('.',',') + ' кг.'
    replace_exact(cell(17,2),'19 407,41 кг.',mass)
    replace_exact(cell(18,2),'05.08.2026',f['start_date'])
    replace_exact(cell(19,2),'09.09.2026',f['end_date'])
    paragraphs = cell(20).findall('w:p',NS)
    if len(paragraphs)!=2 or not text(paragraphs[0]).startswith('Организация, выполнившая рабочую документацию'):
        raise Stage7Error('QUALITY_KM_FIELD_CHANGED')
    replace_exact(paragraphs[1],text(paragraphs[1]),f['km'])
    replace_exact(cell(22),'ООО «ВЗМК» SAMPLE-PROJECT-KMD', f['kmd_company']+' '+f['cipher'])
    replace_exact(cell(24,2),'С255, С345',f['steel'])
    replace_exact(cell(5,1,1),'09.09.2026',f['end_date'])
    banner = cell(10).findall('w:p',NS)
    if len(banner)!=3 or text(banner[1]):
        raise Stage7Error('QUALITY_BANNER_SLOT_CHANGED')
    set_paragraph(banner[1],BANNER,banner=True)
    list_cell(cell(51), [acts_paragraph(manifest,'basis_point13')])
    list_cell(cell(54), [acts_paragraph(manifest,'attachment_point14'), STORAGE])
    structural_check(before,root)
    payload = BytesIO()
    with ZipFile(TEMPLATE) as source, ZipFile(payload,'w') as dest:
        for part in source.infolist():
            dest.writestr(part,E.tostring(root,xml_declaration=True,encoding='UTF-8',standalone=True)
                          if part.filename=='word/document.xml' else source.read(part.filename))
    output = Path(output)
    output.parent.mkdir(parents=True,exist_ok=True)
    with output.open('xb') as stream:
        stream.write(payload.getvalue())
    return output


def validate_quality(path, manifest):
    root=parse(path)
    checks=structural_check(parse(TEMPLATE),root)
    cells=table_cells(root)
    if text(cells[0][50][0]) != acts_paragraph(manifest,'basis_point13'):
        raise Stage7Error('POINT13_REFERENCE_MISMATCH')
    if [text(p) for p in cells[0][53][0].findall('w:p',NS)] != [acts_paragraph(manifest,'attachment_point14'),STORAGE]:
        raise Stage7Error('POINT14_REFERENCE_MISMATCH')
    if text(cells[0][11][2]) != 'Договор №' or text(cells[0][11][3]).strip() != contract_value(manifest['fields']['contract']):
        raise Stage7Error('CONTRACT_NUMBER_MISMATCH')
    if BANNER not in text(root) or 'АЦСТ -118-03329' in text(root):
        raise Stage7Error('QUALITY_TEST_OR_HISTORICAL_TEXT')
    with ZipFile(TEMPLATE) as a, ZipFile(path) as b:
        if a.namelist()!=b.namelist() or any(a.read(n)!=b.read(n) for n in a.namelist() if n!='word/document.xml'):
            raise Stage7Error('QUALITY_OPAQUE_PACKAGE_CHANGED')
    return checks | {'opaque_parts_equal':True,'point13_equal':True,'point14_equal':True,'test_visible':True}
