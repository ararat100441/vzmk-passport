"""Read-only Stage 5 discovery. Writes evidence only beside this script."""
import csv
import hashlib
import json
import os
import re
from datetime import datetime, timezone, timedelta
from pathlib import Path
from zipfile import ZipFile
from xml.etree import ElementTree as ET

OUT = Path(__file__).resolve().parent
APP = OUT.parent.parent
ROOT = APP.parent
LOCAL = timezone(timedelta(hours=4))

def sha(path):
    h = hashlib.sha256()
    with path.open('rb') as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b''):
            h.update(block)
    return h.hexdigest()

def save(name, data):
    (OUT / name).write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding='utf-8')

def walk(root):
    errors, skipped, files = [], [], []
    for current, dirs, names in os.walk(root, followlinks=False, onerror=lambda e: errors.append(str(e))):
        for name in list(dirs):
            p = Path(current) / name
            if p.is_symlink() or p.is_junction():
                dirs.remove(name)
                skipped.append(str(p))
        files.extend(Path(current) / name for name in names)
    return files, errors, skipped

files, errors, skipped = walk(ROOT)
uzk = re.compile(r'(^|[\s_\-])УЗК([\s_.\-]|$)', re.I)
candidates = sorted([p for p in files if uzk.search(p.name) and p.suffix.lower() in ('.docx', '.xlsm', '.xlsx', '.xls')], key=lambda p: str(p).casefold())
protected = [p for p in files if p.is_relative_to(APP) and (
    p.is_relative_to(APP / 'passport_app') or p.is_relative_to(APP / 'profiles_stage4') or
    p.is_relative_to(APP / 'DEV') or p.parent == APP) and 'node_modules' not in p.parts]
protected = sorted(set(protected + candidates), key=str)
before = {str(p): sha(p) for p in protected}
save('preservation_before.json', before)

checkpoint = json.loads((APP / 'DEV' / 'DEV_REGISTRY_CHECKPOINT.json').read_text(encoding='utf-8-sig'))
accepted = json.loads((APP / 'validation' / 'stage4_excel_work_rereview_20261007_39f14b' / 'REVIEW_MANIFEST.json').read_text(encoding='utf-8-sig'))
expected = {str(APP / p): h for p, h in accepted['product_hashes'].items()}
expected.update(accepted['DEV_artifacts'])
expected[checkpoint['path']] = checkpoint['sha256']
baseline = [{'path': p, 'expected': h, 'actual': sha(Path(p)) if Path(p).is_file() else None} for p, h in expected.items()]
for item in baseline:
    item['match'] = item['expected'] == item['actual']
save('accepted_baseline_check.json', baseline)

inventory = []
ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
for p in candidates:
    stat = p.stat()
    record = {'path': str(p), 'size_bytes': stat.st_size,
              'modified_local_utc_plus_04': datetime.fromtimestamp(stat.st_mtime, LOCAL).isoformat(),
              'sha256': before[str(p)], 'extension': p.suffix.lower(),
              'template_selection': 'NOT_RUN_STOP_NO_AUTHORITATIVE_REGISTRY'}
    if p.name.startswith('~$'):
        record.update(type='Office owner/lock file', classification='other: not a DOCX package', purpose='Office temporary lock metadata')
    elif p.suffix.lower() == '.docx':
        try:
            with ZipFile(p) as package:
                document = ET.fromstring(package.read('word/document.xml'))
                text = '\n'.join(''.join(x.itertext()) for x in document.findall('.//w:t', ns))
                paragraphs = [''.join(t.text or '' for t in para.findall('.//w:t', ns)) for para in document.findall('.//w:p', ns)]
                visible = '\n'.join(paragraphs)
                record.update(type='Word OOXML document',
                              purpose='Ultrasonic inspection report' if re.search('ультразвук|ультразвуков|УЗК', visible, re.I) else 'NOT_ESTABLISHED',
                              content_excerpt='\n'.join(x for x in paragraphs if x.strip())[:1100])
                first = bool(re.match(r'^1(?:[_\s/.-]|$)', p.name))
                record['classification'] = ('first-passport project template candidate; selection NOT_RUN' if first else 'existing project report/result; first-passport eligibility NOT_RUN')
                record['classification_basis'] = 'Existing report content; first-passport candidate flag is filename evidence only, not canonical DEV mapping.'
        except Exception as exc:
            record.update(type='unreadable DOCX candidate', classification='other: NOT_ESTABLISHED', purpose='NOT_ESTABLISHED', error=str(exc))
    else:
        record.update(type={'.xls':'Excel binary workbook', '.xlsx':'Excel OOXML workbook', '.xlsm':'Excel macro-enabled OOXML workbook'}[p.suffix.lower()],
                      classification='other: differently named/formatted workbook; authority NOT_ESTABLISHED',
                      purpose='NOT_RUN: content not inspected after discovery STOP')
    inventory.append(record)
save('file_inventory.json', inventory)
columns = ['path','size_bytes','modified_local_utc_plus_04','sha256','type','classification','purpose','template_selection']
with (OUT / 'file_inventory.csv').open('w', encoding='utf-8-sig', newline='') as stream:
    writer = csv.DictWriter(stream, fieldnames=columns, extrasaction='ignore')
    writer.writeheader()
    writer.writerows(inventory)

after = {str(p): sha(p) for p in protected}
save('preservation_after.json', after)
changed = [p for p in before if before[p] != after[p]]
summary = {
    'created_local': datetime.now(LOCAL).isoformat(), 'root': str(ROOT),
    'scan': {'files': len(files), 'errors': errors, 'skipped_reparse_directories': skipped,
             'scope': 'root only; no traversal into junction or symlink targets'},
    'exact_UZK_xlsm': [str(p) for p in candidates if p.name.casefold() == 'узк.xlsm'],
    'UZK_xlsm_candidates': [str(p) for p in candidates if p.suffix.lower() == '.xlsm'],
    'exact_UZK_docx': [str(p) for p in candidates if p.name.casefold() == 'узк.docx'],
    'exact_global_fallback': {'path': str(ROOT / 'Шаблоны' / 'УЗК.docx'), 'exists': (ROOT / 'Шаблоны' / 'УЗК.docx').is_file()},
    'inventory_count': len(inventory),
    'real_docx_count': sum(x['type'] == 'Word OOXML document' for x in inventory),
    'owner_file_count': sum(x['type'] == 'Office owner/lock file' for x in inventory),
    'other_workbooks': [x['path'] for x in inventory if x['extension'] in ('.xls','.xlsx','.xlsm')],
    'accepted_hash_checks': len(baseline), 'accepted_hash_mismatches': [x for x in baseline if not x['match']],
    'protected_count': len(before), 'changed_protected_files': changed,
    'stop': 'Required УЗК.xlsm absent from allowed root; authoritative registry cannot be established. No exact global fallback.',
    'ready_to_implement': False,
    'native_Office_opened': False, 'macro_execution': False, 'source_writes': False,
    'production_registry_opened_or_hashed': False, 'web_access': False,
}
save('audit_summary.json', summary)
print(json.dumps(summary, ensure_ascii=False, indent=2))
