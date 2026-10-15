import json
from pathlib import Path
import sys
from zipfile import ZipFile
from lxml import etree as E

RUN = Path(__file__).resolve().parent
ROOT = RUN.parents[1]
sys.path.insert(0, str(ROOT))
from uzk_stage5.core import Preview, generate, product_paths
from uzk_stage5.inputs import sha

record = json.loads((RUN/'current_preview.json').read_text(encoding='utf8'))
plan = Preview(record['preview'], record['seal'])
baseline = json.loads((RUN/'baseline.json').read_text(encoding='utf8'))
old = baseline['existing_test_docx']
assert sha(old['path']) == old['sha256']
with ZipFile(old['path']) as source:
    root = E.fromstring(source.read('word/document.xml'))
ns = {'w':'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
section = E.tostring(root.find('.//w:sectPr', ns), encoding='unicode')
(RUN/'artifact.md').write_text(
    '# Retained template contract\n\n'
    f'Reference: {old["path"]}\nSHA256: {old["sha256"]}\nNative pages: 2\n'
    f'First-passport generator source: {plan.data["template"]}\n\n'
    'Preserve all package parts, page format, tables, grids, cell and run properties.\n'
    'Preserve first six cells of each mark and all non-title content.\n'
    'Only authorized differences from the accepted TEST: both literal sample cells '
    '(ДНО / годен), title marker and unconfirmed-results line.\n'
    'Use the existing title visible run formatting, no new design presets.\n'
    'Exactly four mark rows, unchanged 33 units, unchanged geometry and equipment.\n'
    'Official results are unavailable; per-mark executor confirmation is a future gate.\n\n'
    'Retained section XML:\n```xml\n'+section+'\n```\n', encoding='utf8')
paths = [*product_paths(), ROOT/'STAGE5_DEV.md', ROOT/'tests/test_stage5.py',
         ROOT/'tests/test_stage5_s5_r1.py', ROOT/'tests/test_stage5_sample_results.py']
binding = {str(path):sha(path) for path in paths}
(RUN/'product_before_generation.json').write_text(json.dumps(binding,ensure_ascii=False,indent=2),encoding='utf8')
output = generate(plan, RUN/'output')
assert sha(old['path']) == old['sha256']
result = {'output':str(output),'sha256':sha(output),'preview_seal':plan.seal,
          'old_test_unchanged':True, 'product_binding':binding, 'mode':'SAMPLE_ONLY',
          'results_confirmed':False}
(RUN/'new_document.json').write_text(json.dumps(result,ensure_ascii=False,indent=2),encoding='utf8')
print(json.dumps(result,ensure_ascii=False,indent=2))
