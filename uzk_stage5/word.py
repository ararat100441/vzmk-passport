"""Patch only document.xml and remove obsolete external PDF relationships."""
from copy import deepcopy
from io import BytesIO
import re
from zipfile import ZipFile
from lxml import etree as E

from profiles_stage4.model import fmt
from .inputs import Stage5Error
from .result_samples import (TEST_LABEL, UNCONFIRMED_LABEL, load_sample_results,
                             verify_sample_sources)

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
R = 'http://schemas.openxmlformats.org/package/2006/relationships'
NS = {'w': W}


def tag(name):
    return '{' + W + '}' + name


def contents(node):
    return ''.join(node.xpath('.//w:t/text()', namespaces=NS))


def replace_text(node, value):
    paragraph = node if node.tag == tag('p') else node.find('w:p', NS)
    if paragraph is None:
        raise Stage5Error('WORD_PARAGRAPH_MISSING')
    # Paragraph-mark formatting can differ from visible text (e.g. bold=false
    # on the actual geometry run). Copy the first visible run, not the mark.
    visible_runs = paragraph.xpath('.//w:r[w:t]', namespaces=NS)
    rpr = visible_runs[0].find('w:rPr', NS) if visible_runs else None
    rpr = deepcopy(rpr) if rpr is not None else None
    for child in list(paragraph):
        if child.tag != tag('pPr'):
            paragraph.remove(child)
    for extra in list(node.findall('w:p', NS))[1:]:
        node.remove(extra)
    run = E.SubElement(paragraph, tag('r'))
    if rpr is not None:
        run.append(rpr)
    for i, line in enumerate(value.split('\n')):
        if i:
            E.SubElement(run, tag('br'))
        t = E.SubElement(run, tag('t'))
        t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
        t.text = line


def write_test_document(data, output):
    sample = load_sample_results(data)
    with ZipFile(data['template']) as source:
        root = E.fromstring(source.read('word/document.xml'))
        body = root.find('w:body', NS)
        tables = body.findall('w:tbl', NS)
        sections = root.findall('.//w:sectPr', NS)
        if len(tables) != 2 or len(sections) != 1:
            raise Stage5Error('WORD_STRUCTURE_CHANGED: ожидаются 2 таблицы / 1 секция')
        if root.xpath('.//w:sdt | .//w:instrText | .//w:fldSimple', namespaces=NS):
            raise Stage5Error('WORD_UNMAPPED_FIELD: требуется отдельное сопоставление полей')
        table = tables[0]
        rows = table.findall('w:tr', NS)
        if len(rows) < 6 or 'Характеристика сварного соединения' not in contents(rows[4]):
            raise Stage5Error('WORD_TABLE_HEADER_CHANGED')
        historical = rows[5:]
        if any(len(r.findall('w:tc', NS)) != 8 for r in historical):
            raise Stage5Error('WORD_MERGED_CELL_MAPPING_CHANGED: требуется 8 логических ячеек')
        if any('L=' not in contents(r.findall('w:tc', NS)[1]) for r in historical):
            raise Stage5Error('WORD_MARK_ROW_MAPPING_CHANGED')
        titles = [p for p in body.findall('w:p', NS) if re.fullmatch(r'Акт № \d+/\d{4} от \d{2}\.\d{2}\.\d{4}', contents(p).strip())]
        if len(titles) != 1:
            raise Stage5Error('WORD_ACT_MAPPING_AMBIGUOUS')
        pp = data['passport']
        replace_text(titles[0], f'Акт № {pp["number"]} от {pp["date"]} ({TEST_LABEL})\n{UNCONFIRMED_LABEL}')
        object_cells = rows[1].findall('w:tc', NS)
        if len(object_cells) != 2 or contents(object_cells[0]).strip() != 'Объект:':
            raise Stage5Error('WORD_OBJECT_MAPPING_CHANGED')
        replace_text(object_cells[1], f'{pp["object"]} ({pp["cipher"]})')
        pattern = deepcopy(historical[0])
        for row in historical:
            table.remove(row)
        for mark in data['marks']:
            row = deepcopy(pattern)
            cells = row.findall('w:tc', NS)
            values = {
                0: mark['mark'],
                1: '\n'.join(f'L={fmt(s["L"])}х{fmt(s["S"])} мм' for s in mark['segments']),
                3: pp['date'],
                4: '\n'.join(fmt(s['frequency_mhz']).replace('.', ',') for s in mark['segments']),
                5: '\n'.join(fmt(s['angle_degrees']) for s in mark['segments']),
                # Literal examples only, covered by the visible unconfirmed label.
                6: sample['graph7'], 7: sample['graph8'],
            }
            for index, value in values.items():
                replace_text(cells[index], value)
            table.append(row)
        for link in root.findall('.//w:hyperlink', NS):
            # No historical local drawing links may remain in a new draft.
            parent = link.getparent()
            index = parent.index(link)
            for child in list(link):
                parent.insert(index, child)
                index += 1
            parent.remove(link)
        rels = E.fromstring(source.read('word/_rels/document.xml.rels'))
        for rel in list(rels):
            if rel.get('Type', '').endswith('/hyperlink'):
                rels.remove(rel)
        edits = {
            'word/document.xml': E.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True),
            'word/_rels/document.xml.rels': E.tostring(rels, xml_declaration=True, encoding='UTF-8', standalone=True),
        }
        payload = BytesIO()
        with ZipFile(payload, 'w') as dest:
            for part in source.infolist():
                dest.writestr(part, edits.get(part.filename, source.read(part.filename)))
        verify_sample_sources(sample)
        # Exclusive create: a race or existing name can never silently overwrite.
        try:
            stream = output.open('xb')
        except FileExistsError as exc:
            raise Stage5Error(f'OUTPUT_EXISTS: {output}') from exc
        try:
            with stream:
                stream.write(payload.getvalue())
        except Exception:
            output.unlink(missing_ok=True)  # this invocation exclusively created it
            raise
