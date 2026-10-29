"""File safety tests on real copies; publication uses verified native Excel outputs."""
import json
from contextlib import nullcontext
from io import BytesIO
from pathlib import Path
from unittest.mock import patch
from xml.etree import ElementTree as ET
from zipfile import ZipFile
from validate_real import copy_case, OUT, BASE, PRODUCTION, SOURCE_VOM, SOURCE_TEMPLATE
from passport_app.config import PassportError
from passport_app.excel import sha256
from passport_app.service import validate, generate
from passport_app.registry import read_registry

checks = []
baseline = {str(p): sha256(p) for p in (PRODUCTION.registry_path, SOURCE_VOM, SOURCE_TEMPLATE)}

def rejected(label, operation, phrase):
    try:
        operation()
    except PassportError as exc:
        assert phrase in str(exc), str(exc)
        checks.append({'check': label, 'pass': True, 'message': str(exc)})
    else:
        raise AssertionError(label + ': expected rejection')

def change_owned_cell(path, coordinate):
    ns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
    output = BytesIO()
    with ZipFile(path) as source, ZipFile(output, 'w') as target:
        for entry in source.infolist():
            payload = source.read(entry.filename)
            if entry.filename == 'xl/worksheets/sheet1.xml':
                sheet = ET.fromstring(payload)
                cell = sheet.find(f'.//{{{ns}}}c[@r="{coordinate}"]')
                assert cell is not None
                for child in list(cell):
                    cell.remove(child)
                cell.set('t', 'inlineStr')
                ET.SubElement(ET.SubElement(cell, f'{{{ns}}}is'), f'{{{ns}}}t').text = 'TEST CHANGED AFTER CHECK'
                payload = ET.tostring(sheet, encoding='utf-8', xml_declaration=True)
            target.writestr(entry, payload)
    path.write_bytes(output.getvalue())

for kind in ('first_registry_row', 'first_template'):
    settings, vom, template = copy_case(kind)
    plan = validate(settings, 'Гараж Казахстан', vom)
    change_owned_cell(settings.registry_path if kind == 'first_registry_row' else template,
                      'C353' if kind == 'first_registry_row' else 'A1')
    modified = sha256(settings.registry_path)
    rejected(kind, lambda: generate(settings, plan), 'после проверки')
    assert sha256(settings.registry_path) == modified and not plan.folder.exists()

# Native generation is independently covered by validate_real and GUI smoke.
# Replay its actual saved files to isolate filesystem publication and recovery.
validated = sorted((BASE / 'validation').glob('*/results.json'))[-1]
native = json.loads(validated.read_text(encoding='utf-8'))['main_result']
saved_registry = Path(native['folder']).parents[3] / 'ВЗМК.xlsx'
saved_appendix = Path(native['appendix'])
assert saved_registry.is_file() and saved_appendix.is_file()
saved_row = read_registry(saved_registry).projects['Гараж Казахстан'][-1]

def staged_registry(excel, path, registry, project, vom, day):
    assert saved_row['date'].date() == day
    path.write_bytes(saved_registry.read_bytes())
    return saved_row['row'], dict(saved_row, date=saved_row['date'].date())

def staged_appendix(excel, path, project, vom, requisites):
    path.write_bytes(saved_appendix.read_bytes())
    return 37

rename = Path.rename
def fail_publication(self, target):
    if self.name.startswith('.pasport-pending-'):
        raise PermissionError('TEST publication failure')
    return rename(self, target)

settings, vom, template = copy_case('publication_failure')
plan = validate(settings, 'Гараж Казахстан', vom)
with patch('passport_app.service.excel_session', lambda: nullcontext(None)), \
     patch('passport_app.service.append_row', staged_registry), \
     patch('passport_app.service.generate_appendix', staged_appendix), \
     patch.object(Path, 'rename', fail_publication):
    rejected('failure_after_registry_commit', lambda: generate(settings, plan), 'Не повторяйте формирование')
assert any(r['number'] == '9/2026' for r in read_registry(settings.registry_path).projects['Гараж Казахстан'])
assert not plan.folder.exists()
rejected('incomplete_transaction_blocks_retry', lambda: validate(settings, 'Гараж Казахстан', vom), 'незавершённая операция')
assert all(sha256(Path(p)) == digest for p, digest in baseline.items())
result = {'pass': True, 'checks': checks, 'production_unchanged': True, 'root': str(OUT),
          'native_generation_source': str(validated),
          'method': 'External edits simulated in copied XLSX; verified native outputs replayed for publication fault; Windows replacement and journal are real.'}
(OUT / 'safety_results.json').write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding='utf-8')
print(json.dumps(result, ensure_ascii=False))
