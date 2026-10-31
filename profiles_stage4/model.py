"""Profile identity, standard-aware lookup and the confirmed per-item UZK rule."""
from collections import Counter, defaultdict
from copy import deepcopy
from decimal import Decimal, InvalidOperation
import hashlib
import json
import re

RULE = 'ДВУТАВР_ПРОКАТНЫЙ_1'
BENT_RULE = 'ГНУТЫЙ_ШВЕЛЛЕР_1'
CLOSED_RULE = 'ЗАМКНУТЫЙ_ПРОФИЛЬ_1'
CHANNEL_RULE = 'ШВЕЛЛЕР_16П_1'
SUPPORTED_RULES = frozenset((RULE, BENT_RULE, CLOSED_RULE, CHANNEL_RULE))
APPROVED = 'ПОДТВЕРЖДЕНО ПОЛЬЗОВАТЕЛЕМ'
NOT_CONFIGURED = 'НЕ НАСТРОЕНО'
GEOMETRY = ('h', 'b', 'b2', 's', 't', 'D', 'r', 'R')
BLOCKING_FIELDS = frozenset(('Тип', 'h', 'b', 'b2', 's', 't', 'D'))
PRIMARY = {
    'Двутавр': ('h', 'b', 's', 't'),
    'Швеллер': ('h', 'b', 's', 't'),
    'Гнутый швеллер': ('h', 'b', 't'),
    'Уголок равнополочный': ('b', 't'),
    'Уголок неравнополочный': ('b', 'b2', 't'),
    'Квадратный профиль': ('h', 'b', 't'),
    'Прямоугольный профиль': ('h', 'b', 't'),
    'Круглая труба': ('D', 's'),
}


class ProfileError(ValueError):
    pass


def number(value, *, optional=False, source_footnote=False):
    if value is None or value == '':
        if optional:
            return None
        raise ProfileError('Missing number')
    if isinstance(value, bool):
        raise ProfileError('Boolean is not a dimension')
    text = str(value).strip().replace('\u00a0', '')
    if source_footnote:
        text = text.removesuffix('*')
    if not re.fullmatch(r'[+-]?\d+(?:[.,]\d+)?', text):
        raise ProfileError(f'Invalid number: {value!r}')
    try:
        val = Decimal(text.replace(',', '.'))
    except InvalidOperation as exc:
        raise ProfileError('Invalid decimal') from exc
    if not val.is_finite():
        raise ProfileError('Non-finite number')
    return int(val) if val == val.to_integral() else float(val)


def fmt(value):
    return format(Decimal(str(value)).normalize(), 'f')


def normalize_profile(text):
    """Recognized family labels only; series letters and asterisks survive."""
    value = str(text).strip().upper().replace('\u00a0', ' ')
    value = re.sub(r'\s+', ' ', value).replace(',', '.')
    compact = re.sub(r'\s+', '', value)
    beam = r'\d+(?:\.\d+)?(?:[А-ЯЁ]+\d*)*'
    candidate = re.sub(r'^(?:ДВУТАВР|I)', '', compact)
    if candidate != compact and re.fullmatch(beam, candidate):
        return candidate  # Preserve every series digit, including a leading zero.
    sizes = r'\d+(?:\.\d+)?(?:[XХ×]\d+(?:\.\d+)?){1,2}'
    # KMD prefixes distinguish an open bent channel from a closed section.
    typed = re.fullmatch(r'ГН\.?(\[\]|\[)(' + sizes + ')', compact)
    if typed:
        parts = [fmt(number(p)) for p in re.split('[XХ×]', typed[2])]
        if typed[1] == '[]':
            if len(parts) == 2:  # Explicit closed-square notation h x t.
                parts.insert(1, parts[0])
            return '×'.join(parts)
        if len(parts) == 3:
            return 'ГН.[' + '×'.join(parts)
    candidate = re.sub(r'^(?:ГНЗ|ПРОФИЛЬНАЯТРУБА|КРУГЛАЯТРУБА|ТРУБА)', '', compact)
    if re.fullmatch(sizes, candidate):
        return '×'.join(fmt(number(p)) for p in re.split('[XХ×]', candidate))
    candidate = re.sub(r'^УГОЛОК', '', compact)
    candidate = candidate.removeprefix('L')
    if (compact.startswith('L') or compact.startswith('УГОЛОК')) and re.fullmatch(sizes, candidate):
        return 'L' + '×'.join(fmt(number(p)) for p in re.split('[XХ×]', candidate))
    if re.fullmatch(beam, compact):
        return compact
    return value  # Opaque/manual names keep word boundaries and significant characters.


def normalize_standard(text):
    value = str(text).strip().upper()
    value = re.sub(r'[‐‑–—−]', '-', value)
    value = re.sub(r'\s*-\s*', '-', value)
    value = re.sub(r'\s+', ' ', value)
    value = re.sub(r'^ГОСТ\s*Р\s*(?=\d)', 'ГОСТ Р ', value)
    value = re.sub(r'^ГОСТ\s*(?=\d)', 'ГОСТ ', value)
    value = re.sub(r'^СТО\s*АСЧМ\s*(?=\d)', 'СТО АСЧМ ', value)
    return value


def validate_geometry(family, geometry):
    if family not in PRIMARY:
        raise ProfileError(f'Unsupported family: {family}')
    result = {k: number(v, optional=True) for k, v in geometry.items() if k in GEOMETRY}
    for field in PRIMARY[family]:
        val = result.get(field)
        if val is None or val <= 0:
            raise ProfileError(f'{family}: missing/invalid {field}')
    for field, val in result.items():
        if val is not None and val < 0:
            raise ProfileError(f'Negative {field}')
    if family in ('Квадратный профиль', 'Прямоугольный профиль'):
        if (result['h'] == result['b']) != (family == 'Квадратный профиль'):
            raise ProfileError('Square/rectangular family does not match dimensions')
        if 2 * result['t'] >= min(result['h'], result['b']):
            raise ProfileError('Invalid tube wall thickness')
    if family == 'Круглая труба' and 2 * result['s'] >= result['D']:
        raise ProfileError('Invalid round wall thickness')
    if family == 'Двутавр' and (2 * result['t'] >= result['h'] or result['s'] >= result['b']):
        raise ProfileError('Invalid I-beam geometry')
    if family == 'Гнутый швеллер' and (2 * result['t'] >= result['h'] or result['t'] >= result['b']):
        raise ProfileError('Invalid bent channel geometry')
    return result


def equal(a, b):
    return abs(Decimal(str(a)) - Decimal(str(b))) <= Decimal('0.000000001')


def signature(row):
    return (row['family'], *(row['geometry'].get(k) for k in PRIMARY[row['family']]))


def frequency_angle(thickness):
    s = number(thickness)
    if s <= 0:
        raise ProfileError('Thickness must be positive')
    return {'frequency_mhz': 5 if s < 12 else 2.5, 'angle_degrees': 70 if s < 12 else 65}


def calculate_uzk(result):
    if result.get('status') not in ('FOUND', 'FOUND_PREFERRED'):
        raise ProfileError('UZK requires resolved lookup')
    row = result['profile']
    rule = row['uzk_rule']
    if rule not in SUPPORTED_RULES:
        return {'status': NOT_CONFIGURED, 'rows': []}
    g = validate_geometry(row['family'], row['geometry'])
    if rule == RULE and row['family'] == 'Двутавр':
        values = [(Decimal(str(g['h'])) - 2 * Decimal(str(g['t'])), g['s'], 'стенка'),
                  (2 * Decimal(str(g['b'])), g['t'], 'полки')]
    elif rule == CHANNEL_RULE and row['family'] == 'Швеллер' and row['key'] == '16П' and row['standard'] == 'ГОСТ 8240-97':
        values = [(Decimal(str(g['h'])), g['s'], 'стенка'), (2 * Decimal(str(g['b'])), g['t'], 'полки')]
    elif (rule == BENT_RULE and row['family'] == 'Гнутый швеллер') or (rule == CLOSED_RULE and row['family'] in ('Квадратный профиль', 'Прямоугольный профиль')):
        values = [(2 * (Decimal(str(g['h'])) + Decimal(str(g['b']))), g['t'], 'периметр')]
    else:
        raise ProfileError('UZK rule does not match profile identity/type')
    groups = {}
    for length, thickness, part in values:
        key = Decimal(str(thickness))
        item = groups.setdefault(key, {'L': 0, 'S': thickness, 'parts': [], **frequency_angle(thickness)})
        item['L'] = number(str(Decimal(str(item['L'])) + length))
        item['parts'].append(part)
    return {'status': 'CALCULATED', 'rule': rule, 'scope': 'ONE_ITEM', 'rows': list(groups.values())}


def approved_extension_rule(row, members, state):
    """Explicit persisted scopes only; an approval never resolves a conflict."""
    matches = set()
    for approval in state.get('uzk_approvals', []):
        if approval.get('status') != APPROVED or not approval.get('basis') or not approval.get('date'):
            continue
        if normalize_standard(approval.get('standard', '')) != row['standard']:
            continue
        if row['family'] not in approval.get('families', []):
            continue
        key = approval.get('key')
        if key is not None and normalize_profile(key) != row['key']:
            continue
        rule = approval.get('rule')
        if rule == CHANNEL_RULE:
            valid = row['key'] == '16П' and row['standard'] == 'ГОСТ 8240-97' and row['family'] == 'Швеллер' and key is not None
        elif rule == BENT_RULE:
            valid = row['family'] == 'Гнутый швеллер' and key is not None and all(m['source'] == 'USER' and m.get('verification', {}).get('kind') == 'KMD' for m in members)
        elif rule == CLOSED_RULE:
            valid = row['family'] in ('Квадратный профиль', 'Прямоугольный профиль') and row['standard'] == 'ГОСТ 30245-2003' and all(m['source'] != 'USER' or m.get('verification', {}).get('kind') == 'KMD' for m in members)
        else:
            valid = False
        if valid:
            validate_geometry(row['family'], row['geometry'])
            matches.add(rule)
    return next(iter(matches)) if len(matches) == 1 else NOT_CONFIGURED


def default_user_state():
    return {'version': 1, 'revision': 0, 'manual_rows': [], 'resolutions': {}, 'preferences': [
        {'key': '35Б1', 'standard': 'ГОСТ Р 57837-2017', 'basis': 'USER_CONFIRMED_UZK_328x6_348x9',
         'status': 'ПОДТВЕРЖДЕНО ПОЛЬЗОВАТЕЛЕМ', 'date': '2026-10-06'}]}


class Catalog:
    def __init__(self, records, sources=None, user_state=None):
        self.raw_records = deepcopy(records)
        self.sources = deepcopy(sources or [])
        self.user_state = deepcopy(user_state if user_state is not None else default_user_state())
        all_records = self.raw_records + deepcopy(self.user_state['manual_rows'])
        groups = defaultdict(list)
        for record in all_records:
            record['key'] = normalize_profile(record['key'])
            record['standard'] = normalize_standard(record['standard'])
            groups[(record['key'], record['standard'])].append(record)
        self.rows, self.conflicts = [], []
        for (key, standard), members in sorted(groups.items()):
            families = {r['family'] for r in members}
            all_values = defaultdict(list)
            for r in members:
                for field, value in {**r['geometry'], **r.get('quality', {})}.items():
                    if value is not None:
                        all_values[field].append(value)
            differences = [f for f, vals in all_values.items() if any(not equal(vals[0], v) for v in vals[1:])]
            if len(families) != 1:
                differences.append('Тип')
            fingerprint = hashlib.sha256(json.dumps(members, ensure_ascii=False, sort_keys=True).encode()).hexdigest()
            conflict_id = 'CF-' + hashlib.sha256((key + '\0' + standard).encode()).hexdigest()[:16]
            resolution = self.user_state['resolutions'].get(conflict_id)
            chosen = None
            if resolution and resolution.get('fingerprint') == fingerprint:
                chosen = next((r for r in members if r['record_id'] == resolution.get('record_id')), None)
            conflicting = bool(differences)
            blocking = bool(BLOCKING_FIELDS.intersection(differences))
            classification = ('BLOCKING' if blocking else 'NON_BLOCKING') if conflicting else ''
            row = {'profile': members[0]['display'], 'key': key, 'standard': standard,
                   'family': members[0]['family'], 'geometry': {}, 'quality': {},
                   'provenance': members, 'conflict_id': conflict_id if conflicting else '',
                   'collision': len(families) != 1,
                   'disagreement_class': classification,
                   'unresolved_fields': differences[:] if conflicting and not chosen else [],
                   'notes': sorted({note for r in members for note in r.get('notes', [])}),
                   'uzk_rule': NOT_CONFIGURED}
            if conflicting:
                self.conflicts.append({'id': conflict_id, 'key': key, 'standard': standard,
                    'profile': row['profile'], 'fields': differences, 'members': members,
                    'classification': classification,
                    'fingerprint': fingerprint, 'resolution': resolution or None,
                    'status': 'РАЗРЕШЕНО ПОЛЬЗОВАТЕЛЕМ' if chosen else
                              'КОНФЛИКТ' if blocking else 'СПРАВОЧНЫЕ ДАННЫЕ НЕ РАЗРЕШЕНЫ'})
            if blocking and not chosen:
                row['status'] = 'КОНФЛИКТ'
                row['notes'].append('Геометрия не выбрана; требуется явное решение по Conflict ID')
            else:
                if chosen:
                    row['geometry'], row['quality'] = deepcopy(chosen['geometry']), deepcopy(chosen.get('quality', {}))
                    row['status'] = 'РАЗРЕШЕНО ПОЛЬЗОВАТЕЛЕМ'
                    rule_members = [chosen]
                else:
                    # Only unanimous values become canonical. Raw members retain every disputed value.
                    row['geometry'] = {f: vals[0] for f, vals in all_values.items() if f in GEOMETRY and f not in differences}
                    row['quality'] = {f: vals[0] for f, vals in all_values.items() if f not in GEOMETRY and f not in differences}
                    manual = any(r['source'] == 'USER' for r in members)
                    row['status'] = ('РУЧНОЙ' if manual else 'ПРОВЕРЕНО_ДВУМЯ_ИСТОЧНИКАМИ'
                                     if len({r['source'] for r in members}) > 1 else 'ИЗ ИСТОЧНИКА')
                    if conflicting:
                        row['status'] = 'СОГЛАСОВАНО_С_ЗАМЕЧАНИЯМИ'
                        row['notes'].append('NON_BLOCKING: основные размеры согласованы; unresolved: ' + ', '.join(differences))
                    rule_members = members
                if row['family'] == 'Двутавр' and all(r.get('uzk_confirmed', False) for r in rule_members):
                    row['uzk_rule'] = RULE
                else:
                    row['uzk_rule'] = approved_extension_rule(row, rule_members, self.user_state)
            self.rows.append(row)
        self.by_key = defaultdict(list)
        for row in self.rows:
            self.by_key[row['key']].append(row)

    def lookup(self, profile_text, standard=None):
        key = normalize_profile(profile_text)
        rows = self.by_key.get(key, [])
        if not rows and re.fullmatch(r'\[\d+(?:\.\d+)?[А-ЯЁ]*', key):
            # Lookup-only fallback. Do not strip arbitrary brackets or hide collisions.
            candidates = self.by_key.get(key[1:], [])
            if candidates and (len(candidates) != 1 or candidates[0]['family'] != 'Швеллер' or candidates[0]['collision']):
                return {'status': 'AMBIGUOUS', 'key': key, 'candidates': deepcopy(candidates)}
            if candidates:
                key, rows = key[1:], candidates
        if standard is not None:
            rows = [r for r in rows if r['standard'] == normalize_standard(standard)]
        if not rows:
            return {'status': 'NOT_FOUND', 'key': key, 'candidates': []}
        if any(r['collision'] and r['status'] != 'КОНФЛИКТ' for r in rows):
            return {'status': 'AMBIGUOUS', 'key': key, 'candidates': deepcopy(rows)}
        def found(status, chosen, matching=None):
            return {'status': status, 'key': key, 'profile': deepcopy(chosen),
                    'candidates': deepcopy(rows), 'equivalent_standards': [r['standard'] for r in (matching or [chosen])]}
        if standard is not None or len(rows) == 1:
            return found('CONFLICT' if rows[0]['status'] == 'КОНФЛИКТ' else 'FOUND', rows[0])
        prefs = [p for p in self.user_state['preferences'] if normalize_profile(p['key']) == key
                 and p['status'] == 'ПОДТВЕРЖДЕНО ПОЛЬЗОВАТЕЛЕМ']
        preferred = {normalize_standard(p['standard']) for p in prefs}
        # A conflicting preferred identity is never rescued by a different standard.
        if len(preferred) == 1:
            chosen = next((r for r in rows if r['standard'] in preferred), None)
            if chosen and chosen['status'] == 'КОНФЛИКТ':
                return found('CONFLICT', chosen)
        if any(r['status'] == 'КОНФЛИКТ' for r in rows):
            if len(preferred) == 1:
                chosen = next((r for r in rows if r['standard'] in preferred), None)
                if chosen and chosen['status'] != 'КОНФЛИКТ':
                    return found('FOUND_PREFERRED', chosen)
            return {'status': 'CONFLICT', 'key': key, 'candidates': deepcopy(rows)}
        if len({signature(r) for r in rows}) == 1:
            chosen = deepcopy(rows[0])
            chosen['provenance'] = [p for r in rows for p in r['provenance']]
            # UZK permission must hold for every equivalent candidate.
            if len({r['uzk_rule'] for r in rows}) != 1:
                chosen['uzk_rule'] = NOT_CONFIGURED
            for f in ('r', 'R'):
                vals = [r['geometry'].get(f) for r in rows if r['geometry'].get(f) is not None]
                if vals and any(not equal(vals[0], v) for v in vals[1:]):
                    chosen['geometry'].pop(f, None)
                    chosen['notes'].append(f'{f}: разные значения по НД; см. candidates')
            return found('FOUND', chosen, rows)
        if len(preferred) == 1:
            chosen = next((r for r in rows if r['standard'] in preferred), None)
            if chosen:
                return found('FOUND_PREFERRED', chosen)
        return {'status': 'AMBIGUOUS', 'key': key, 'candidates': deepcopy(rows)}

    def counts(self):
        return {'canonical': len(self.rows), 'source_raw': len(self.raw_records),
                'raw_by_source': dict(Counter(r['source'] for r in self.raw_records)),
                'by_family': dict(Counter(r['family'] for r in self.rows)),
                'by_standard': dict(Counter(r['standard'] for r in self.rows)),
                'by_status': dict(Counter(r['status'] for r in self.rows)),
                'conflicts': len(self.conflicts),
                'blocking_disagreements': sum(c['classification'] == 'BLOCKING' for c in self.conflicts),
                'non_blocking_disagreements': sum(c['classification'] == 'NON_BLOCKING' for c in self.conflicts),
                'unresolved_conflicts': sum(c['status'] == 'КОНФЛИКТ' for c in self.conflicts),
                'ambiguous_keys': sum(self.lookup(k)['status'] == 'AMBIGUOUS' for k in self.by_key),
                'manual_rows': len(self.user_state['manual_rows']),
                'uzk_enabled': sum(r['uzk_rule'] in SUPPORTED_RULES for r in self.rows),
                'uzk_by_rule': dict(Counter(r['uzk_rule'] for r in self.rows if r['uzk_rule'] in SUPPORTED_RULES)),
                'uzk_by_family': dict(Counter(r['family'] for r in self.rows if r['uzk_rule'] in SUPPORTED_RULES)),
                'missing_optional_radius': sum(r['geometry'].get('R') is None and r['geometry'].get('r') is None for r in self.rows),
                'missing_R': sum(r['geometry'].get('R') is None for r in self.rows),
                'missing_r': sum(r['geometry'].get('r') is None for r in self.rows)}
