"""Strict synthetic evidence contract; TEST cannot become OFFICIAL."""
from collections import defaultdict
from pathlib import Path
import re
import unicodedata
from zipfile import ZipFile
import xml.etree.ElementTree as ET
from .storage import Stop, sha, read, within, digest
from .native import sheet_part


def key(text):
    if not isinstance(text, str) or not text.strip():
        raise Stop('IDENTITY_MISSING')
    return ' '.join(unicodedata.normalize('NFKC', text).casefold().split())


def number(value):
    if not isinstance(value, str) or not re.fullmatch(r'[1-9]\d*/20\d{2}', value):
        raise Stop('PASSPORT_NUMBER_INVALID')
    return value


def integer(value, zero=False):
    if isinstance(value, bool) or not isinstance(value, (int, float)) or int(value) != value or value < (0 if zero else 1):
        raise Stop('QUANTITY_INVALID')
    return int(value)


def ref(value, root, bindings):
    p = within(value['path'], root/'inputs')
    if not re.fullmatch(r'[a-f0-9]{64}', str(value.get('sha256', ''))) or sha(p) != value['sha256']:
        raise Stop('INPUT_SHA256_CHANGED: '+str(p))
    bindings[str(p)] = value['sha256']
    return read(p)


def membership(data, identity):
    if any(data.get(k) != v for k,v in identity.items()):
        raise Stop('EVIDENCE_MEMBERSHIP_MISMATCH')
    if data.get('test_only') is not True:
        raise Stop('SYNTHETIC_TEST_EVIDENCE_REQUIRED')


def registry_rows(path, project, marks):
    # Direct read-only OOXML parser: never openpyxl-save a macro workbook.
    ns = {'s': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
    with ZipFile(path) as z:
        strings = []
        if 'xl/sharedStrings.xml' in z.namelist():
            strings = [''.join(n.itertext()) for n in ET.fromstring(z.read('xl/sharedStrings.xml')).findall('s:si',ns)]
        sheet = ET.fromstring(z.read(sheet_part(path)))
        values, formulas = {}, set()
        for c in sheet.findall('.//s:c', ns):
            coord = c.get('r')
            if c.find('s:f', ns) is not None:
                formulas.add(coord)
            v = c.find('s:v', ns)
            if c.get('t') == 's':
                value = strings[int(v.text)] if v is not None else None
            elif c.get('t') == 'inlineStr':
                value = ''.join(c.find('s:is', ns).itertext())
            else:
                value = float(v.text) if v is not None and v.text is not None else None
            values[coord] = value
    def header(v):
        return re.sub(r'\s+', '', str(v)).casefold()
    if tuple(header(values.get(c+'1')) for c in 'ABC') != tuple(header(v) for v in ('Заказ','Марка элемента','Кол-во (шт)')):
        raise Stop('UZK_HEADER_CHANGED')
    found = defaultdict(list)
    for coord, value in values.items():
        if coord.startswith('A') and coord[1:].isdigit() and int(coord[1:]) > 1 and value is not None and key(str(value)) == key(project):
            r = int(coord[1:])
            mark = values.get('B'+str(r))
            if mark is not None and key(str(mark)) in marks:
                raw = values.get('C'+str(r))
                if 'C'+str(r) in formulas:
                    raise Stop('UZK_BALANCE_FORMULA_FORBIDDEN')
                found[key(str(mark))].append({'row':r,'raw_before':raw,
                    'before': 1 if raw is None else integer(raw, zero=True), 'mark':str(mark)})
    result = {}
    for mark in marks:
        if len(found[mark]) != 1:
            raise Stop('UZK_MARK_MISSING_OR_DUPLICATE: '+mark)
        result[mark] = found[mark][0]
    return result


def prepare(root, request_path):
    request_path = within(request_path, root/'inputs')
    q = read(request_path)
    if q.get('schema') != 'STAGE8_CLOSE_REQUEST_V1' or q.get('mode') != 'ISOLATED_TEST' or q.get('official_issue') is not False:
        raise Stop('OFFICIAL_CLOSURE_FORBIDDEN')
    if q.get('passport') == '9/2026':
        raise Stop('PASSPORT_9_2026_DEV_TEST_FORBIDDEN')
    identity = {k:q[k] for k in ('project','passport','shipment_id')}
    key(identity['project']);number(identity['passport']);key(identity['shipment_id'])
    bindings = {str(request_path):sha(request_path)}
    shipment = ref(q['shipment'],root,bindings)
    requirements = ref(q['requirements'],root,bindings)
    control = ref(q['control'],root,bindings)
    for source in (shipment, requirements, control):
        membership(source, identity)
    if shipment.get('schema') != 'STAGE8_SYNTHETIC_SHIPMENT_V1' or requirements.get('schema') != 'STAGE8_REQUIREMENTS_V1' or control.get('schema') != 'STAGE8_CONFIRMED_TEST_CONTROL_V1':
        raise Stop('EVIDENCE_SCHEMA_INVALID')
    if type(requirements.get('uzk_required')) is not bool or type(requirements.get('akz_required')) is not bool:
        raise Stop('EXPLICIT_REQUIREMENTS_MISSING')
    if control.get('status') != 'CONFIRMED_TEST' or control.get('all_required_control_confirmed') is not True:
        raise Stop('CONTROL_NOT_CONFIRMED')
    amounts = defaultdict(int)
    if not shipment.get('rows'):
        raise Stop('SHIPMENT_EMPTY')
    for row in shipment['rows']:
        amounts[key(row['mark'])] += integer(row['quantity'])
    required = {}
    for row in requirements['uzk']:
        m = key(row['mark'])
        if m in required or m not in amounts:
            raise Stop('UZK_REQUIREMENT_MEMBERSHIP_INVALID')
        required[m] = integer(row['quantity'])
        if required[m] > amounts[m]:
            raise Stop('UZK_EXCEEDS_SHIPMENT')
    confirmed = {}
    for row in control['uzk']:
        m = key(row['mark'])
        if m in confirmed or row.get('confirmed') is not True:
            raise Stop('UZK_CONFIRMATION_INVALID')
        confirmed[m] = integer(row['quantity'])
    if confirmed != required or bool(required) != requirements['uzk_required']:
        raise Stop('UZK_CONTROL_INCOMPLETE_OR_REQUIREMENT_CONFLICT')
    required_kinds = {'QUALITY','APPENDIX_1','VIK'}
    if requirements['uzk_required']:
        required_kinds.add('UZK')
    if requirements['akz_required']:
        required_kinds.update(('AKZ','AKZ_APPENDIX'))
    docs = {}
    for entry in q['documents']:
        doc = ref(entry,root,bindings)
        membership(doc, identity)
        kind = entry['kind']
        if (doc.get('schema') != 'STAGE8_SYNTHETIC_DOCUMENT_V1' or kind != doc.get('kind')
                or kind in docs or doc.get('status') != 'CONFIRMED_TEST'
                or doc.get('results_confirmed') is not True):
            raise Stop('DOCUMENT_NOT_CONFIRMED_OR_DUPLICATE')
        docs[kind] = entry
    missing = required_kinds - set(docs)
    if missing:
        raise Stop('MANDATORY_DOCUMENTS_MISSING: '+','.join(sorted(missing)))
    if ('UZK' in docs) != requirements['uzk_required']:
        raise Stop('UZK_DOCUMENT_REQUIREMENT_CONFLICT')
    registry = root/'registry/УЗК_TEST.xlsm'
    balances = registry_rows(registry, q['project'], set(required)) if required else {}
    deductions = []
    for m, count in required.items():
        d = balances[m] | {'quantity':count, 'after':balances[m]['before']-count}
        if d['after'] < 0:
            raise Stop('INSUFFICIENT_UZK_BALANCE: '+d['mark'])
        deductions.append(d)
    return {'identity':identity,'key':digest({'project':key(q['project']),'passport':number(q['passport'])}),
        'mode':'ISOLATED_TEST','official_issue':False,'documents':q['documents'],
        'required_documents':sorted(required_kinds),'control_status':'CONFIRMED_TEST',
        'uzk_required':requirements['uzk_required'],'controlled_units':sum(required.values()),
        'deductions':deductions,'bindings':bindings,'before_sha256':sha(registry),
        'registry':str(registry),'request_path':str(request_path),'request_sha256':bindings[str(request_path)]}
