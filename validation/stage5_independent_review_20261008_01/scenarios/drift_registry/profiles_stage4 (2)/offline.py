"""Pinned accepted snapshots. No URL fetch, no execution of audit scripts/macros."""
from collections import defaultdict
import hashlib
import json
from pathlib import Path
import re
import socket
from contextlib import contextmanager
from unittest.mock import patch
from .model import ProfileError, fmt, normalize_profile, number, validate_geometry, GEOMETRY

ROOT = Path(__file__).resolve().parents[1]
A_REL = 'validation/stage4_km_kmd_audit_20261006_204347'
B_REL = 'validation/stage4_ptoflow_audit_20261006_215239'
PINS = {A_REL: '5ee23d78efa923def1c49153a414e7881a813c48e7f8390602dbf7cac95fb1e0',
        B_REL: '33f604f3f7dbeb495d21ecf480c77c358ea7cb5fbfb50ee5b1da8b061e2fc485'}
ALLOW_A = {'Двутавр': 'СТО АСЧМ 20-93', 'Швеллер': 'ГОСТ 8240-97',
           'Уголок равнополочный': 'ГОСТ 8509-93', 'Уголок неравнополочный': 'ГОСТ 8510-86'}
ALLOW_B = {'ГОСТ 26020-83': 'Двутавр', 'ГОСТ Р 57837-2017': 'Двутавр', 'ГОСТ 8239-89': 'Двутавр',
           'ГОСТ 8240-97': 'Швеллер', 'ГОСТ 8509-93': 'Уголок равнополочный',
           'ГОСТ 8510-86': 'Уголок неравнополочный', 'ГОСТ 30245-2003': 'Профильная труба',
           'ГОСТ 10704-91': 'Круглая труба', 'ГОСТ 8732-2025': 'Круглая труба'}


def sha(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def read_json(path):
    return json.loads(Path(path).read_text(encoding='utf-8-sig'))


@contextmanager
def no_network():
    """Fail closed even if a future Python dependency attempts a fetch."""
    attempts = []
    def denied(*args, **kwargs):
        attempts.append('network call blocked')
        raise ProfileError('Stage 4 is offline: network disabled')
    with patch.object(socket.socket, 'connect', denied), patch.object(socket.socket, 'connect_ex', denied), \
         patch.object(socket.socket, 'sendto', denied), patch.object(socket, 'getaddrinfo', denied), \
         patch.object(socket, 'create_connection', denied):
        yield attempts


def checked(root, relative):
    folder = (root / relative).resolve()
    manifest_path = folder / 'AUDIT_MANIFEST.json'
    if sha(manifest_path) != PINS[relative]:
        raise ProfileError('Accepted manifest hash changed: ' + relative)
    manifest = read_json(manifest_path)
    for name, meta in manifest['evidence_files'].items():
        path = (folder / name.replace('\\', '/')).resolve()
        if not path.is_relative_to(folder) or not path.is_file() or sha(path) != meta['sha256']:
            raise ProfileError('Accepted evidence hash changed: ' + name)
    return folder, manifest


def ingest_a(root=ROOT):
    folder, manifest = checked(root, A_REL)
    records = read_json(folder / 'parsed_source_records.json')
    sources = []
    download = manifest['download']
    sources.append({'id': 'A-ZIP', 'source': 'KM-KMD', 'type': 'ZIP', 'family': '4 семейства', 'standard': '',
        'reference': download['exact_download_url'], 'date': download['downloaded_at_utc'],
        'snapshot': str((folder / 'sortament.zip').relative_to(root)), 'sha256': download['sha256'],
        'count': len(records), 'status': 'ПРИНЯТ', 'notes': download['source_page']})
    # Verify audit values against the actual read-only XLS cell origins.
    import xlrd
    books = {}
    for workbook in sorted({r['workbook'] for r in records}):
        path = folder / 'extracted' / workbook
        books[workbook] = xlrd.open_workbook(str(path), on_demand=True)
        rows = [r for r in records if r['workbook'] == workbook]
        sources.append({'id': 'A-' + workbook, 'source': 'KM-KMD', 'type': 'XLS/BIFF8 read-only',
            'family': rows[0]['family'], 'standard': rows[0]['standard'], 'reference': str(path.relative_to(root)),
            'date': download['downloaded_at_utc'], 'snapshot': str(path.relative_to(root)), 'sha256': sha(path),
            'count': len(rows), 'status': 'ПРИНЯТ С QUALITY NOTES', 'notes': 'ZIP A-ZIP; inherited cells retained in provenance'})
    result = []
    try:
        for i, record in enumerate(records):
            family, standard = record['family'], record['standard']
            if ALLOW_A.get(family) != standard:
                raise ProfileError('Source A contains an unaccepted dataset')
            sheet = books[record['workbook']].sheet_by_name(record['sheet'])
            for field, origin in record['provenance'].items():
                rr, cc = origin['origin']
                actual = sheet.cell_value(rr - 1, cc - 1)
                actual = None if actual == '' else actual
                if actual != record['effective'][field]:
                    raise ProfileError(f'XLS origin disagrees with audit: {record["workbook"]}:{rr}:{cc}:{field}')
            p = record['parsed']
            for field, val in record['effective'].items():
                if field != 'source_profile' and number(val, optional=True, source_footnote=True) != p[field]:
                    raise ProfileError('Source A numeric evidence mismatch')
            geometry = validate_geometry(family, {k: p[k] for k in GEOMETRY if k in p})
            raw_name = str(record['effective']['source_profile'])
            if family.startswith('Уголок'):
                # Printed corner number is in cm. It defines identity; erroneous dimensions stay untouched.
                sizes = raw_name.removesuffix('*').split('/')
                sides = [number(str(number(v) * 10)) for v in sizes]
                key = 'L' + '×'.join(map(fmt, sides + [geometry['t']]))
                display = raw_name  # Original printed designation remains visible.
            else:
                key, display = normalize_profile(raw_name), raw_name
            notes = []
            if any(geometry.get(f) is None for f in ('R', 'r')):
                notes.append('Отсутствующий R/r не вычислялся')
            if family.startswith('Уголок') and any(not (a == geometry.get(f)) for a, f in zip(sides, ('b', 'b2'))):
                notes.append('Размеры отличаются от исходного номера уголка; identity из номера, геометрия без исправления')
            if family == 'Двутавр' and key in ('18Б2', '50Б2'):
                notes.append('SOURCE_A_AUDIT: подозрительная площадь; другой НД не исправляет исходное значение')
            result.append({'record_id': f'A-{i+1:04d}', 'source_id': 'A-' + record['workbook'], 'source': 'KM-KMD',
                'display': display, 'key': key, 'family': family, 'standard': standard, 'geometry': geometry,
                'quality': {'area_cm2': p.get('area_cm2'), 'mass_kg_m': p.get('mass_kg_m')},
                'raw': record['raw'], 'effective': record['effective'], 'date': download['downloaded_at_utc'],
                'evidence': {'file': str((folder/'parsed_source_records.json').relative_to(root)), 'record': i+1,
                    'workbook': record['workbook'], 'sheet': record['sheet'], 'row': record['row'],
                    'start_column': record['start_column'], 'cells': record['provenance']},
                'notes': notes, 'uzk_confirmed': family == 'Двутавр'})
    finally:
        for book in books.values():
            book.release_resources()
    return result, sources


def flight_rows(path):
    """JSON decode only; never execute JavaScript from cache."""
    from lxml import html
    tree = html.fromstring(path.read_bytes())
    decoder = json.JSONDecoder()
    payload = ''
    for script in tree.xpath('//script[not(@src)]'):
        code = script.text or ''
        for hit in re.finditer(re.escape('self.__next_f.push('), code):
            array, _ = decoder.raw_decode(code[hit.end():])
            if len(array) > 1 and isinstance(array[1], str):
                payload += array[1]
    datasets = []
    for hit in re.finditer(r'"rows":', payload):
        try:
            rows, _ = decoder.raw_decode(payload[hit.end():])
        except ValueError:
            continue
        if isinstance(rows, list) and rows and isinstance(rows[0], dict) and 'size' in rows[0]:
            datasets.append(rows)
    if len(datasets) != 1:
        raise ProfileError('Group cache must contain exactly one rows dataset')
    return datasets[0]


def ingest_b(root=ROOT):
    folder, _ = checked(root, B_REL)
    evidence = read_json(folder / 'group_source_evidence.json')
    pages = {p['url']: p for p in read_json(folder / 'pages.json') if p['kind'] == 'group'}
    grouped = defaultdict(list)
    for r in evidence:
        if ALLOW_B.get(r['standard']) != r['family'] or not r['primary_geometry_complete']:
            raise ProfileError('Unaccepted PTOflow group evidence')
        grouped[r['url']].append(r)
    result, sources = [], []
    for url, rows in grouped.items():
        page = pages[url]
        path = folder / page['local_path'].replace('\\', '/')
        if page['http_status'] != 200 or sha(path) != page['sha256']:
            raise ProfileError('Invalid cached group')
        raw_rows = flight_rows(path)
        if len(raw_rows) != len(rows):
            raise ProfileError('Group evidence count differs from HTML')
        source_id = 'B-' + url.rsplit('/', 1)[1]
        sources.append({'id': source_id, 'source': 'PTOflow', 'type': 'cached HTML / embedded numeric group',
            'family': rows[0]['family'], 'standard': rows[0]['standard'], 'reference': url,
            'date': page['timestamp_utc'], 'snapshot': str(path.relative_to(root)), 'sha256': page['sha256'],
            'count': len(rows), 'status': 'ПРИНЯТ GROUP', 'notes': 'Numeric sort values; category/detail rows excluded'})
        for r in rows:
            rownum = r['embedded_row']
            raw = raw_rows[rownum - 1]
            if raw != r['raw']:
                raise ProfileError('Raw group evidence differs from cached HTML')
            numeric = {f: number(v['sort']) for f, v in raw.items() if isinstance(v, dict) and 'sort' in v}
            family = r['family']
            label = raw['size']['label'] if isinstance(raw['size'], dict) else raw['size']
            if family == 'Уголок неравнополочный':
                geometry = {'b': numeric['B'], 'b2': numeric['b'], 't': numeric['t']}
            elif family == 'Круглая труба':
                geometry = {'D': numeric['d'], 's': numeric['s']}
            elif family == 'Профильная труба':
                key = normalize_profile(label)
                if not re.fullmatch(r'\d+(?:\.\d+)?×\d+(?:\.\d+)?×\d+(?:\.\d+)?', key):
                    raise ProfileError('Unproved profile-tube format')
                h, b, t = map(number, key.split('×'))
                if t != numeric['t']:
                    raise ProfileError('Tube label thickness differs from numeric group')
                geometry = {'h': h, 'b': b, 't': t, 'R': numeric.get('R')}
                family = 'Квадратный профиль' if h == b else 'Прямоугольный профиль'
            else:
                geometry = {f: numeric[f] for f in PRIMARY_FIELDS(family)}
            geometry = validate_geometry(family, geometry)
            for field, val in geometry.items():
                if r['parsed'].get(field) != val:
                    raise ProfileError('Parsed dimensions disagree with proved audit values')
            if family == 'Круглая труба':
                key = fmt(geometry['D']) + '×' + fmt(geometry['s'])
            else:
                key = normalize_profile(label)
            notes = []
            rounding = [f for f, val in raw.items() if isinstance(val, dict) and 'sort' in val
                        and f in ('h', 'b', 's', 't', 'd', 'B', 'R')
                        and number(val.get('display'), optional=True) != numeric[f]]
            if rounding:
                notes.append('Display rounding ignored: ' + ','.join(rounding))
            if family == 'Круглая труба' and normalize_profile(label) != key:
                notes.append('Исходное обозначение округлено; точный ключ из numeric D/s')
            result.append({'record_id': source_id + ':' + str(rownum), 'source_id': source_id, 'source': 'PTOflow',
                'display': label, 'key': key, 'family': family, 'standard': r['standard'], 'geometry': geometry,
                'quality': {'area_cm2': numeric.get('area'), 'mass_kg_m': numeric.get('mass')},
                'raw': raw, 'date': r['retrieval_date'], 'evidence': {'file': str((folder/'group_source_evidence.json').relative_to(root)),
                    'url': url, 'cache': str(path.relative_to(root)), 'sha256': page['sha256'], 'embedded_row': rownum},
                'notes': notes, 'display_rounding': rounding, 'uzk_confirmed': family == 'Двутавр'})
    return result, sources


def PRIMARY_FIELDS(family):
    from .model import PRIMARY
    return PRIMARY[family]


def load_sources(root=ROOT):
    with no_network() as attempts:
        a, sa = ingest_a(root)
        b, sb = ingest_b(root)
    if attempts:
        raise ProfileError('Network attempted during ingestion')
    return a + b, sa + sb
