"""Verify identities in extracted evidence; no product implementation or XLSX writes."""
import hashlib
import json
import re
from pathlib import Path

OUT = Path(__file__).resolve().parent
books = json.loads((OUT / 'first_passport_workbooks.json').read_text(encoding='utf-8'))
registry = json.loads((OUT / 'registry_first_passport.json').read_text(encoding='utf-8'))
canonical = registry['project']['first_number']
act_pattern = re.compile(r'\s*(?P<number>[1-9]\d*/\d{4}(?:-[1-9]\d*)?)\s+от\s*')
app_pattern = re.compile(r'Приложение\s+№\s*(?P<own>[1-9]\d*/\d{4}(?:-[1-9]\d*)?)\s+'
                         r'к\s+акту\s+АКЗ\s+№\s*(?P<number>[1-9]\d*/\d{4}(?:-[1-9]\d*)?)\s+'
                         r'от\s+(?P<date>\d{2}\.\d{2}\.\d{4})\s+г\.')
acts = []
apps = []
for book in books:
    if hashlib.sha256(Path(book['path']).read_bytes()).hexdigest() != book['sha256']:
        raise RuntimeError('Workbook changed since extraction')
    for sheet in book['sheets']:
        cells = sheet['cells']
        if cells.get('B3') == 'СЕРТИФИКАТ КАЧЕСТВА АНТИКОРРОЗИОННОЙ ОБРАБОТКИ':
            match = act_pattern.fullmatch(str(cells.get('E4', '')))
            if not match:
                raise RuntimeError('Unparseable act identity: STOP')
            acts.append({'book': book, 'sheet': sheet, 'number': match['number']})
        header = str(cells.get('A1', ''))
        if 'к акту АКЗ' in header:
            match = app_pattern.fullmatch(header)
            if not match:
                raise RuntimeError('Unparseable appendix identity: STOP')
            apps.append({'book': book, 'sheet': sheet, **match.groupdict()})

rows = []
for variant in range(4):
    expected = canonical + (f'-{variant}' if variant else '')
    aa = [a for a in acts if a['number'] == expected]
    pp = [p for p in apps if p['number'] == expected and p['own'] == expected]
    status = 'AMBIGUOUS' if len(aa) > 1 or len(pp) > 1 else (
             'AVAILABLE' if len(aa) == len(pp) == 1 else 'MISSING')
    row = {'akz': variant, 'expected_template_internal_number': expected,
           'act_count': len(aa), 'appendix_count': len(pp), 'status': status,
           'act': None, 'appendix': None, 'layers': None,
           'material': None, 'ral': None, 'required_dry_microns': None}
    if status == 'AVAILABLE':
        a, p = aa[0], pp[0]
        cells, formulas = a['sheet']['cells'], a['sheet']['formulas']
        act_day = cells['H4'][:10]
        yy, mm, dd = act_day.split('-')
        if p['date'] != f'{dd}.{mm}.{yy}':
            raise RuntimeError('Act/appendix date conflict: STOP')
        if cells['E7'] != registry['project']['object_name']:
            raise RuntimeError('Project object conflict: STOP')
        if formulas.get('G8') != '=E4' or formulas.get('I8') != '=H4':
            raise RuntimeError('Act reference to own appendix not proven: STOP')
        if cells.get('G8') != cells['E4'] or cells.get('I8') != cells['H4']:
            raise RuntimeError('Cached reference conflict: STOP')
        layer_headers = {k: v for k, v in cells.items()
                         if re.fullmatch(r'[E-L]17', k)}
        if layer_headers != {'E17': 1}:
            raise RuntimeError('Layer interpretation needs separate review: STOP')
        for role, source in [('act', a), ('appendix', p)]:
            row[role] = {k: source['book'][k]
                         for k in ('path', 'name', 'extension', 'sha256')}
        row.update(layers=1, material=cells['E18'], ral=cells['E21'],
                   required_dry_microns=cells['E25'], date=act_day,
                   identity_proof={'act_number_cell': '4!E4', 'act_date_cell': '4!H4',
                       'act_appendix_number_formula': '4!G8 = E4',
                       'act_appendix_date_formula': '4!I8 = H4',
                       'appendix_header_cell': 'Приложение АКЗ!A1',
                       'appendix_header': p['sheet']['cells']['A1'],
                       'act_layer_headers': layer_headers},
                   filename_only_selection='STOP: filename number 4 conflicts with internal 1/2026',
                   binding_method='explicit project + variant -> exact paths/SHA; verify canonical internal identities and dates',
                   product_binding_installed=False)
    rows.append(row)

result = {'scope': 'first passport only; variants are current-passport suffixes, not previous passports',
          'rows': rows, 'template_readiness': 'PARTIAL',
          'variant_0_test_template_sufficiency': 'YES',
          'variant_0_condition': 'use exact verified pair; never infer passport number from filename 4',
          'implementation_started': False}
(OUT / 'verified_template_bindings.json').write_text(
    json.dumps(result, ensure_ascii=False, indent=2) + '\n', encoding='utf-8')
print(json.dumps({'statuses': {r['akz']: r['status'] for r in rows},
                  'act_candidates': len(acts), 'appendix_candidates': len(apps),
                  'implementation_started': False}))
