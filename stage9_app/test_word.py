"""Stage 9 only: visibly label draft VIK and clarify UZK selection in output copies.

The accepted Stage 2/5 generators and project templates remain untouched.
Original generated files are retained within the individual validation run.
"""
from hashlib import sha256
from io import BytesIO
from pathlib import Path
from uuid import uuid4
import os
from zipfile import ZIP_DEFLATED, ZipFile

from lxml import etree as E
from stage8_closure import storage as S
from uzk_stage5.word import contents, replace_text, tag

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
PKG = 'http://schemas.openxmlformats.org/package/2006/relationships'
CT = 'http://schemas.openxmlformats.org/package/2006/content-types'
NS = {'w': W}
TEST_VIK = 'ТЕСТ / ЧЕРНОВИК — ВИК НЕ ПОДТВЕРЖДЕН'


def _xml(value):
    return E.tostring(value, xml_declaration=True, encoding='UTF-8', standalone=True)


def _zip_rewrite(source_bytes, edits, additions=None):
    additions = additions or {}
    with ZipFile(BytesIO(source_bytes)) as old:
        names = set(old.namelist())
        if not set(edits) <= names or names & set(additions):
            raise ValueError('TEST_WORD_ZIP_STRUCTURE_CHANGED')
        with BytesIO() as memory:
            with ZipFile(memory, 'w', compression=ZIP_DEFLATED) as new:
                for info in old.infolist():
                    new.writestr(info, edits.get(info.filename, old.read(info.filename)))
                for name, body in additions.items():
                    new.writestr(name, body)
            return memory.getvalue()


def _test_vik(source_bytes):
    with ZipFile(BytesIO(source_bytes)) as book:
        root = E.fromstring(book.read('word/document.xml'))
        rels = E.fromstring(book.read('word/_rels/document.xml.rels'))
        types = E.fromstring(book.read('[Content_Types].xml'))
        if 'word/header1.xml' in book.namelist() or len(root.findall('.//w:sectPr', NS)) != 1:
            raise ValueError('VIK_TEST_HEADER_LAYOUT_UNEXPECTED')
        if root.xpath('.//w:headerReference', namespaces=NS):
            raise ValueError('VIK_TEST_EXISTING_HEADER_STOP')
        titles = [''.join(p.xpath('.//w:t/text()', namespaces=NS))
                  for p in root.findall('.//w:p', NS)]
        if sum(t.startswith('АКТ № ') for t in titles) != 2:
            raise ValueError('VIK_TEST_TWO_ACTS_NOT_PROVEN')
        current_ids = {r.get('Id') for r in rels}
        id_ = next('rId'+str(i) for i in range(1,1000) if 'rId'+str(i) not in current_ids)
        new_relation = E.SubElement(rels, '{'+PKG+'}Relationship')
        new_relation.set('Id', id_)
        new_relation.set('Type', R+'/header')
        new_relation.set('Target', 'header1.xml')
        new_type = E.SubElement(types, '{'+CT+'}Override')
        new_type.set('PartName', '/word/header1.xml')
        new_type.set('ContentType', 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml')
        sect = root.find('.//w:sectPr', NS)
        ref = E.Element(tag('headerReference'))
        ref.set(tag('type'), 'default')
        ref.set('{'+R+'}id', id_)
        sect.insert(0, ref)
        header = E.Element(tag('hdr'), nsmap={'w': W})
        p = E.SubElement(header, tag('p'))
        ppr = E.SubElement(p, tag('pPr'))
        E.SubElement(ppr, tag('jc')).set(tag('val'), 'center')
        E.SubElement(ppr, tag('spacing')).set(tag('after'), '0')
        run = E.SubElement(p, tag('r'))
        rpr = E.SubElement(run, tag('rPr'))
        E.SubElement(rpr, tag('b'))
        E.SubElement(rpr, tag('color')).set(tag('val'), 'C00000')
        E.SubElement(rpr, tag('sz')).set(tag('val'), '16')
        E.SubElement(run, tag('t')).text = TEST_VIK
    output = _zip_rewrite(source_bytes,
                          {'word/document.xml': _xml(root),
                           'word/_rels/document.xml.rels': _xml(rels),
                           '[Content_Types].xml': _xml(types)},
                          {'word/header1.xml': _xml(header)})
    with ZipFile(BytesIO(output)) as z:
        written = E.fromstring(z.read('word/header1.xml'))
        if ''.join(written.xpath('.//w:t/text()', namespaces=NS)) != TEST_VIK:
            raise ValueError('VIK_TEST_HEADER_WRITE_FAILED')
    return output


def _test_uzk(source_bytes, marks, expected_quantity):
    from uzk_stage5.result_samples import TEST_LABEL, UNCONFIRMED_LABEL
    if not marks or len({m['mark'] for m in marks}) != len(marks):
        raise ValueError('UZK_TEST_MARKS_EMPTY_OR_DUPLICATE')
    controlled = 0
    with ZipFile(BytesIO(source_bytes)) as z:
        root = E.fromstring(z.read('word/document.xml'))
        body = root.find('w:body', NS)
        tables = body.findall('w:tbl', NS)
        if len(tables) != 2:
            raise ValueError('UZK_TEST_SIGNATURE_TABLE_MISSING')
        rows = tables[0].findall('w:tr', NS)
        if len(rows) != 5+len(marks) or len(rows[4].findall('w:tc', NS)) != 8:
            raise ValueError('UZK_TEST_ROWS_MAPPING_CHANGED')
        if 'Характеристика сварного соединения' not in contents(rows[4]):
            raise ValueError('UZK_TEST_TABLE_HEADER_CHANGED')
        text = ''.join(root.xpath('.//w:t/text()', namespaces=NS))
        if TEST_LABEL not in text or UNCONFIRMED_LABEL not in text or 'Контроль выполнил' not in text:
            raise ValueError('UZK_TEST_UNCONFIRMED_LABEL_OR_SIGNATURE_MISSING')
        header = rows[4].findall('w:tc', NS)[0]
        replace_text(header, 'Марка\nУЗК/ВОМ, шт.')
        for i, mark in enumerate(marks):
            cells = rows[5+i].findall('w:tc', NS)
            if len(cells) != 8 or contents(cells[0]) != mark['mark']:
                raise ValueError('UZK_TEST_MARK_ORDER_CHANGED')
            q, shipped = int(mark['Q']), int(mark['vom_quantity'])
            if q <= 0 or shipped < q:
                raise ValueError('UZK_TEST_QUANTITY_INVALID')
            controlled += q
            replace_text(cells[0], f'{mark["mark"]}\n{q}/{shipped}')
        if controlled != expected_quantity:
            raise ValueError('UZK_TEST_CONTROLLED_QUANTITY_CONFLICT')
        # Repeat the requisites and column headings when long tables span pages.
        # Compact row spacing prevents the signature from being orphaned.
        for row in rows[:5]:
            trpr = row.find('w:trPr', NS)
            if trpr is None:
                trpr = E.Element(tag('trPr'))
                row.insert(0, trpr)
            if trpr.find('w:tblHeader', NS) is None:
                E.SubElement(trpr, tag('tblHeader')).set(tag('val'), '1')
        for row in rows[5:]:
            for para in row.findall('.//w:p', NS):
                ppr = para.find('w:pPr', NS)
                if ppr is None:
                    ppr = E.Element(tag('pPr'))
                    para.insert(0, ppr)
                spacing = ppr.find('w:spacing', NS)
                if spacing is None:
                    spacing = E.SubElement(ppr, tag('spacing'))
                for key,value in [('before','0'),('after','0'),('line','200'),('lineRule','auto')]:
                    spacing.set(tag(key), value)
        sections = root.findall('.//w:sectPr', NS)
        if len(sections) != 1:
            raise ValueError('UZK_TEST_SECTION_CHANGED')
        # Four-mark Kazakhstan sample was rendered with its signature on page 1.
        # Larger tables retain source margins pending native Word QA.
        if len(marks) <= 4:
            margins = sections[0].find('w:pgMar', NS)
            if margins is None:
                raise ValueError('UZK_TEST_MARGINS_MISSING')
            margins.set(tag('top'), '500')
            margins.set(tag('bottom'), '500')
    output = _zip_rewrite(source_bytes, {'word/document.xml': _xml(root)})
    with ZipFile(BytesIO(output)) as z:
        check = E.fromstring(z.read('word/document.xml'))
        actual = check.findall('w:body/w:tbl', NS)[0].findall('w:tr', NS)[5:]
        if ([contents(r.findall('w:tc', NS)[0]) for r in actual] !=
            [f'{m["mark"]}{int(m["Q"])}/{int(m["vom_quantity"])}' for m in marks] or
            any(len(r.findall('w:tc', NS)[0].xpath('.//w:br', namespaces=NS)) != 1 for r in actual)):
            raise ValueError('UZK_TEST_QUANTITIES_NOT_SAVED')
    return output


def _publish(path, run, updated_bytes, *, audit_name):
    path = S.no_links(path)
    run = S.no_links(run)
    if not (run.name.startswith('s9_') and path.is_relative_to(run/'d'/'p')
            and path.suffix.lower() == '.docx'):
        raise ValueError('TEST_WORD_OUTPUT_OUTSIDE_OWNED_RUN')
    original = path.read_bytes()
    archive = run/audit_name
    S.exclusive_bytes(archive, original)
    if S.sha(archive) != sha256(original).hexdigest():
        raise ValueError('TEST_WORD_BASELINE_BACKUP_CHANGED')
    pending = path.with_name('.stage9-word-'+uuid4().hex+'.docx')
    try:
        with pending.open('xb') as f:
            f.write(updated_bytes)
            f.flush(); os.fsync(f.fileno())
        os.replace(pending, path)
    finally:
        if pending.exists():
            pending.unlink()
    return {'baseline':str(archive),'baseline_sha256':sha256(original).hexdigest(),
            'output':str(path),'output_sha256':S.sha(path),'official_issue':False}


def mark_vik(path, run):
    path = S.no_links(path)
    return _publish(path,run,_test_vik(path.read_bytes()),audit_name='VIK_BEFORE_TEST_MARKER.docx')


def annotate_uzk(path,run,marks,expected_quantity):
    path = S.no_links(path)
    return _publish(path,run,_test_uzk(path.read_bytes(),marks,expected_quantity),
                    audit_name='UZK_BEFORE_TEST_LAYOUT.docx')
