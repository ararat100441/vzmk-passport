"""Stage 9 DEV/TEST DOCX wrappers only; never modify the accepted source templates."""
from copy import deepcopy
from io import BytesIO
from pathlib import Path
from zipfile import ZipFile
from lxml import etree as E

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
P = 'http://schemas.openxmlformats.org/package/2006/relationships'
CT = 'http://schemas.openxmlformats.org/package/2006/content-types'
NS = {'w': W}
MARK = 'ТЕСТ / ЧЕРНОВИК — ВИК НЕ ПОДТВЕРЖДЕН'


def w(name):
    return '{' + W + '}' + name


def _rewrite(path, edited, appended=None):
    """Repackage the private TEST copy, preserving unchanged package parts."""
    with ZipFile(path) as source:
        original = {item.filename: source.read(item.filename) for item in source.infolist()}
        members = list(source.infolist())
    if any(name not in original for name in edited):
        raise ValueError('DRAFT_WORD_PART_MISSING')
    if any(name in original for name in (appended or {})):
        raise ValueError('DRAFT_WORD_PART_CONFLICT')
    output = BytesIO()
    with ZipFile(output, 'w') as target:
        for member in members:
            target.writestr(member, edited.get(member.filename, original[member.filename]))
        for name, value in (appended or {}).items():
            target.writestr(name, value)
    return output.getvalue()


def _save_new(path, payload):
    """Exclusive output; neither an existing document nor a source is overwritten."""
    with Path(path).open('xb') as output:
        output.write(payload)


def stamp_vik(src, dest):
    """Add a visible DEV/TEST warning to the common running header of both VIK acts."""
    src, dest = Path(src), Path(dest)
    with ZipFile(src) as source:
        document = E.fromstring(source.read('word/document.xml'))
        relationships = E.fromstring(source.read('word/_rels/document.xml.rels'))
        types = E.fromstring(source.read('[Content_Types].xml'))
        if any(name.startswith('word/header') for name in source.namelist()):
            raise ValueError('VIK_ALREADY_HAS_HEADER_STOP')
        sections = document.xpath('.//w:sectPr', namespaces=NS)
        if len(sections) != 1 or sections[0].find(w('headerReference')) is not None:
            raise ValueError('VIK_SECTION_HEADER_AMBIGUOUS')
        acts = document.xpath('.//w:tbl[.//w:t[contains(.,"АКТ №")]]', namespaces=NS)
        if len(acts) != 2:
            raise ValueError('VIK_TWO_ACTS_NOT_FOUND')
        used = {relation.get('Id') for relation in relationships}
        index = 1
        while 'rId' + str(index) in used:
            index += 1
        rid = 'rId' + str(index)
        rel = E.SubElement(relationships, '{' + P + '}Relationship')
        rel.attrib.update({'Id': rid, 'Type': R + '/header', 'Target': 'header1.xml'})
        header = E.Element(w('hdr'), nsmap={'w': W})
        para = E.SubElement(header, w('p'))
        props = E.SubElement(para, w('pPr'))
        E.SubElement(props, w('jc')).set(w('val'), 'center')
        E.SubElement(props, w('spacing')).set(w('after'), '0')
        run = E.SubElement(para, w('r'))
        font = E.SubElement(run, w('rPr'))
        E.SubElement(font, w('b'))
        E.SubElement(font, w('color')).set(w('val'), 'C00000')
        E.SubElement(font, w('sz')).set(w('val'), '16')
        E.SubElement(run, w('t')).text = MARK
        ref = E.Element(w('headerReference'))
        ref.attrib.update({w('type'): 'default', '{' + R + '}id': rid})
        sections[0].insert(0, ref)
        override = E.SubElement(types, '{' + CT + '}Override')
        override.attrib.update({
            'PartName': '/word/header1.xml',
            'ContentType': 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml',
        })
        xml = lambda value: E.tostring(value, xml_declaration=True, encoding='UTF-8', standalone=True)
        payload = _rewrite(src, {
            'word/document.xml': xml(document),
            'word/_rels/document.xml.rels': xml(relationships),
            '[Content_Types].xml': xml(types),
        }, {'word/header1.xml': xml(header)})
    _save_new(dest, payload)
    with ZipFile(dest) as check:
        if MARK.encode('utf-8') not in check.read('word/header1.xml'):
            raise ValueError('VIK_TEST_BANNER_MISSING')
    return str(dest)


def _replace_cell(cell, value):
    paragraph = cell.find(w('p'))
    if paragraph is None:
        raise ValueError('UZK_CELL_PARAGRAPH_MISSING')
    runs = paragraph.xpath('.//w:r[w:t]', namespaces=NS)
    style = runs[0].find(w('rPr')) if runs else None
    style = deepcopy(style) if style is not None else None
    for child in list(paragraph):
        if child.tag != w('pPr'):
            paragraph.remove(child)
    for child in list(cell.findall(w('p')))[1:]:
        cell.remove(child)
    run = E.SubElement(paragraph, w('r'))
    if style is not None:
        run.append(style)
    for index, line in enumerate(value.split('\n')):
        if index:
            E.SubElement(run, w('br'))
        text = E.SubElement(run, w('t'))
        text.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
        text.text = line


def polish_uzk(src, dest, marks):
    """Show selected quantity per mark and improve pagination in small TEST acts.

    The ratio Q/VOМ means selected pieces / visible shipment pieces. This does
    not confirm that an ultrasonic inspection actually took place.
    """
    src, dest = Path(src), Path(dest)
    if not marks:
        raise ValueError('UZK_NO_SELECTED_MARKS')
    with ZipFile(src) as source:
        root = E.fromstring(source.read('word/document.xml'))
        body = root.find(w('body'))
        tables = body.findall(w('tbl'))
        if len(tables) != 2:
            raise ValueError('UZK_TABLES_CHANGED')
        rows = tables[0].findall(w('tr'))
        if len(rows) != 5 + len(marks):
            raise ValueError('UZK_MARK_ROW_COUNT_CHANGED')
        headings = rows[4].findall(w('tc'))
        if len(headings) != 8:
            raise ValueError('UZK_HEADER_COLUMNS_CHANGED')
        _replace_cell(headings[0], 'Марка\nУЗК/ВОМ')
        total = 0
        for mark, row in zip(marks, rows[5:]):
            cells = row.findall(w('tc'))
            if len(cells) != 8:
                raise ValueError('UZK_ROW_COLUMNS_CHANGED')
            original_mark = ''.join(cells[0].xpath('.//w:t/text()', namespaces=NS))
            if original_mark != mark['mark']:
                raise ValueError('UZK_MARK_SEQUENCE_CHANGED: ' + repr(original_mark))
            selected, shipped = int(mark['Q']), int(mark['vom_quantity'])
            if not 1 <= selected <= shipped:
                raise ValueError('UZK_SELECTION_OUT_OF_RANGE: ' + original_mark)
            total += selected
            _replace_cell(cells[0], f'{original_mark}\n{selected}/{shipped}')
        # Four-row Kazakhstan sample: make the signatures fit on the same page
        # without shrinking text. Large shipments retain the source page margins.
        if len(marks) <= 4:
            for row in rows[5:]:
                for paragraph in row.xpath('.//w:p', namespaces=NS):
                    props = paragraph.find(w('pPr'))
                    if props is None:
                        props = E.Element(w('pPr'))
                        paragraph.insert(0, props)
                    spacing = props.find(w('spacing'))
                    if spacing is None:
                        spacing = E.Element(w('spacing'))
                        props.insert(0, spacing)
                    for key, value in [('before', '0'), ('after', '0'),
                                       ('line', '200'), ('lineRule', 'auto')]:
                        spacing.set(w(key), value)
            section = root.find('.//w:sectPr', NS)
            if section is None or section.find(w('pgMar')) is None:
                raise ValueError('UZK_SECTION_MARGINS_MISSING')
            margin = section.find(w('pgMar'))
            for key in ('top', 'bottom'):
                margin.set(w(key), '500')
        # Word repeats consecutive top table rows on continuation pages.
        for row in rows[:5]:
            props = row.find(w('trPr'))
            if props is None:
                props = E.Element(w('trPr'))
                row.insert(0, props)
            if props.find(w('tblHeader')) is None:
                E.SubElement(props, w('tblHeader')).set(w('val'), '1')
        payload = _rewrite(src, {
            'word/document.xml': E.tostring(root, xml_declaration=True,
                                             encoding='UTF-8', standalone=True),
        })
    _save_new(dest, payload)
    return {'output': str(dest), 'selected_units': total, 'mark_rows': len(marks)}
