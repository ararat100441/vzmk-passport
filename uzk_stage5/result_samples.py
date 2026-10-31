"""Literal historical examples for TEST only; never confirmed inspection results."""
from hashlib import sha256
from io import BytesIO
from pathlib import Path
from zipfile import BadZipFile, ZipFile

from lxml import etree as E

from .inputs import Stage5Error, key, sha

NS = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
TEST_LABEL = 'ТЕСТ / ОБРАЗЕЦ ЗАПОЛНЕНИЯ'
UNCONFIRMED_LABEL = 'РЕЗУЛЬТАТЫ КОНТРОЛЯ НЕ ПОДТВЕРЖДЕНЫ'

# The complete bounded sample, independently extracted from ten project acts.
# A changed/missing act requires a new review, never an automatic replacement.
ACCEPTED_SOURCES = (
    ('Паспорт/1_2026 - УЗК.docx', 'e175f4b6fdde8fd49a7071e963e0ce228b4b7ab7e43a998374664b6b2825810c', 11),
    ('Паспорт/2 от 15.09.2026/2_2026 - УЗК.docx', '0d03ecdaaaa51c4ed66ee98e359a8806c6fa3431eff3210942d4c72de28f9bcf', 6),
    ('Паспорт/3 от 18.09.2026/3_2026 - УЗК.docx', '548a1ad01d5be1647ed8edc4fb5ba4b98dacbb2dde1842eca0c2aeb5cec4f689', 3),
    ('Паспорт/4 от 24.09.2026/4_2026 - УЗК.docx', '5ca853c58b24f329f9524eafcc03532d506eb6fcc737137cbb1200989efdd4db', 6),
    ('Паспорт/5 от 25.09.2026/5_2026 - УЗК.docx', '31ca17c6266e3e65f156894d5b8885affcd77989293d020d0f5020ef060b8921', 4),
    ('Паспорт/6 от 30.09.2026/6_2026 - УЗК.docx', '6c648011f62999a70afd3ad0a200e71b28d292b8c2868bb789b7d4620589c3e0', 2),
    ('Паспорт/7 от 01.10.2026/7_2026 - УЗК.docx', '3511d2b4ed57f6556011d41e7c01f1c6bc0631f26688fbe346c74b3a86e228c9', 10),
    ('Паспорт/8 от 02.10.2026/8_2026 - УЗК.docx', 'acf29e3346f844a434c632b91676af32368461ce1be1e4876e939ada3b85b41a', 7),
    ('Паспорт/9 от 06.10.2026/9_2026 - УЗК.docx', '534069ff76f6ad964ebe1f0929e9f2dbd21d83acf951c87a4fef829150e93100', 4),
    ('Паспорт/10 от 07.10.2026/10_2026 - УЗК.docx', '5e021c93c85348336aa200aa75ff1c5af119f27e3b92963905aeb49788353cbe', 3),
)


def cell_text(cell):
    """Preserve raw case, spaces, tabs, line breaks and paragraph boundaries."""
    paragraphs = []
    for paragraph in cell.findall('w:p', NS):
        text = []
        for item in paragraph.iter():
            name = E.QName(item).localname
            if item.tag == '{' + NS['w'] + '}t':
                text.append(item.text or '')
            elif item.tag in ('{' + NS['w'] + '}br', '{' + NS['w'] + '}cr'):
                text.append('\n')
            elif name == 'tab' and item.tag.startswith('{' + NS['w'] + '}'):
                text.append('\t')
        paragraphs.append(''.join(text))
    return '\n'.join(paragraphs)


def extract_rows(payload, source):
    try:
        with ZipFile(BytesIO(payload)) as archive:
            root = E.fromstring(archive.read('word/document.xml'))
        matches = []
        for table in root.findall('w:body/w:tbl', NS):
            rows = table.findall('w:tr', NS)
            for index, row in enumerate(rows):
                cells = row.findall('w:tc', NS)
                if (len(cells) == 8
                        and 'Обнаруженные дефекты' in cell_text(cells[6]).replace('\n', ' ')
                        and 'Заключение' in cell_text(cells[7])):
                    matches.append((rows, index))
        if len(matches) != 1:
            raise Stage5Error(f'SAMPLE_TABLE_AMBIGUOUS: {source}')
        rows, header = matches[0]
        result = []
        for index, row in enumerate(rows[header + 1:], header + 2):
            cells = row.findall('w:tc', NS)
            if len(cells) != 8:
                raise Stage5Error(f'SAMPLE_ROW_MAPPING_CHANGED: {source}, строка {index}')
            values = [cell_text(cell) for cell in cells]
            if any(value.strip() for value in values):
                result.append({'source': str(source), 'row': index, 'mark': values[0],
                               'graph7': values[6], 'graph8': values[7]})
        return result
    except (BadZipFile, KeyError, E.XMLSyntaxError) as exc:
        raise Stage5Error(f'SAMPLE_INVALID_DOCUMENT: {source}') from exc


def assess_rows(rows):
    if not rows or any(not row['mark'].strip() or not row['graph7'].strip()
                       or not row['graph8'].strip() for row in rows):
        raise Stage5Error('SAMPLE_NOT_PROVEN: отсутствует заполненный достоверный образец')
    pairs = {(row['graph7'], row['graph8']) for row in rows}
    if len(pairs) != 1:
        raise Stage5Error('SAMPLE_AMBIGUOUS: несовместимые формулировки; требуется решение пользователя')
    return next(iter(pairs))


def load_sample_results(data):
    if data.get('kind') != 'TEST/DRAFT':
        raise Stage5Error('OFFICIAL_RESULTS_NOT_IMPLEMENTED: требуется отдельное подтверждение '
                          'фактических результатов ответственным исполнителем по каждой марке')
    if key(data['passport']['project']) != key('Гараж Казахстан'):
        raise Stage5Error('SAMPLE_PROJECT_NOT_PROVEN: образец принят только для Гараж Казахстан')
    project = Path(data['project_root']).resolve()
    records, fingerprints = [], {}
    for relative, expected, count in ACCEPTED_SOURCES:
        source = (project / relative).resolve()
        if not source.is_relative_to(project):
            raise Stage5Error('SAMPLE_PATH_OUTSIDE_PROJECT')
        try:
            payload = source.read_bytes()
        except OSError as exc:
            raise Stage5Error(f'SAMPLE_SOURCE_MISSING: {source}') from exc
        if sha256(payload).hexdigest() != expected:
            raise Stage5Error(f'SAMPLE_SOURCE_CHANGED: {source}. Требуется проверка образца.')
        rows = extract_rows(payload, source)
        if len(rows) != count:
            raise Stage5Error(f'SAMPLE_ROW_COUNT_CHANGED: {source}')
        records.extend(rows)
        fingerprints[str(source)] = expected
    graph7, graph8 = assess_rows(records)
    if len(fingerprints) != 10 or len(records) != 56 or (graph7, graph8) != ('ДНО', 'годен'):
        raise Stage5Error('SAMPLE_NOT_PROVEN: образец не совпадает с принятыми десятью актами')
    return {'mode': 'SAMPLE_ONLY', 'confirmed': False, 'graph7': graph7, 'graph8': graph8,
            'row_count': len(records), 'fingerprints': fingerprints}


def verify_sample_sources(sample):
    for path, expected in sample['fingerprints'].items():
        if not Path(path).is_file() or sha(path) != expected:
            raise Stage5Error(f'SAMPLE_SOURCE_CHANGED: {path}')
