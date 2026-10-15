import hashlib
import json
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
from datetime import datetime, timezone

OUT = Path(__file__).resolve().parent
APP = OUT.parent.parent

def read(p):
    return json.loads(Path(p).read_text(encoding='utf-8-sig'))

def save(name, data):
    (OUT / name).write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding='utf-8')

def sha(p):
    return hashlib.sha256(Path(p).read_bytes()).hexdigest()

ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
      'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'}
links = []
for item in read(OUT / 'word_inputs.json'):
    with zipfile.ZipFile(item['path']) as z:
        root = ET.fromstring(z.read('word/document.xml'))
        rels = {r.attrib['Id']: r.attrib for r in ET.fromstring(z.read('word/_rels/document.xml.rels'))}
        found = []
        for el in root.findall('.//w:hyperlink', ns):
            rel = rels.get(el.attrib.get('{'+ns['r']+'}id'), {})
            found.append({'display': ''.join(el.itertext()), 'relationship': rel})
        links.append({'path': item['path'], 'sha256': sha(item['path']), 'hyperlinks': found,
                      'explicit_xml_field_codes': len(root.findall('.//w:instrText', ns)),
                      'native_note': 'Word exposes w:hyperlink as type 88 HYPERLINK fields; see native_word_fields.json'})
save('word_hyperlinks.json', links)

# Supersede the initial wrongly scaled diagnostic raster with the six native
# landscape pages already inspected. No original document is edited.
bad = (OUT / 'garage_first_page.png').resolve()
assert bad.parent == OUT.resolve()
if bad.exists():
    bad.unlink()
native = read(OUT / 'native_word.json')
native[0]['visual_evidence'] = ['garage_first_1.png', 'garage_first_2.png']
native[0]['visual_note'] = 'Initial wrong-aspect diagnostic raster removed; native_render.json is authoritative for all six pages.'
save('native_word.json', native)
script = OUT / 'native_word.ps1'
text = script.read_text(encoding='utf-8-sig')
start = text.index('            if($results.Count -eq 0){')
end = text.index('\n        } catch {', start)
text = text[:start] + '            # Visual evidence is produced separately by native_render.ps1.\n' + text[end:]
script.write_text(text, encoding='utf-8-sig')

report_builder = OUT / 'complete_audit.py'
text = report_builder.read_text(encoding='utf-8')
text = text.replace("'no fields or header/footer parts'", "'no PAGE fields or header/footer parts; mark hyperlinks separately documented'")
needle = '## 7. DOCUMENT FIELD MAPPING'
addition = ('В выбранном первом DOCX 11 mark hyperlinks: native Word Fields.Count=11, все type88 HYPERLINK; '
            'это relationships w:hyperlink, поэтому XML instrText/fldChar=0 не означает отсутствие hyperlink fields. '
            'Ссылки ведут к старым локальным PDF сборок (например К1-5.pdf); они не открывались. '
            'Во втором гаражном и ТЭМПО examples native Fields.Count=0. PAGE fields отсутствуют. '
            'Evidence: word_hyperlinks.json и native_word_fields.json.\n\n')
text = text.replace(needle, addition + needle)
text = text.replace("['Profile / quantity separate columns'", "['T1 data C1 hyperlink','11 historical PDF links, e.g. К1-5.pdf','No approved current drawing-link source; replace old hyperlink with plain current mark unless separately approved','DYNAMIC / REMOVE STALE LINK'],\n['Profile / quantity separate columns'")
text = text.replace('No new profile/quantity columns.', 'No new profile/quantity columns. Historical mark hyperlink targets must not survive row cloning; use plain current mark until an authoritative drawing-link source is separately approved.')
text = text.replace('Paragraphs XML', 'Paragraphs XML')
report_builder.write_text(text, encoding='utf-8')

before = read(OUT / 'preservation_before.json')
after = {p: sha(p) for p in before}
assert before == after, 'STOP protected file drift'
save('preservation_after.json', after)
old = APP / 'validation' / 'stage5_uzk_audit_20261007_stop'
old_manifest = read(old / 'EVIDENCE_MANIFEST.json')
old_changes = [name for name, expected in old_manifest['evidence_sha256'].items() if sha(old / name) != expected]
assert not old_changes, 'STOP previous evidence drift'
save('final_integrity_check.json', {'checked_at_utc': datetime.now(timezone.utc).isoformat(),
     'protected_files': len(before), 'protected_changes': [],
     'previous_evidence_files': len(old_manifest['evidence_sha256']), 'previous_evidence_changes': [],
     'native_hyperlink_count': len(links[0]['hyperlinks']),
     'authoritative_visual_evidence': 'native_render.json; six landscape PNG pages',
     'working_mark_catalog_exists': (APP.parent / 'Гараж казахстан' / 'Справочник марок.xlsx').exists()})
print(json.dumps({'protected': len(before), 'previous_evidence': len(old_manifest['evidence_sha256']),
                  'hyperlinks': len(links[0]['hyperlinks']), 'changed': 0}))
