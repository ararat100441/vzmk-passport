"""Fail closed on any drift; write review manifest after every other evidence file."""
import ast
from datetime import datetime, timezone
import difflib
import json
from pathlib import Path
import sys

RUN=Path(__file__).resolve().parent
ROOT=RUN.parents[1]
sys.path.insert(0,str(ROOT))
from uzk_stage5.core import product_paths
from uzk_stage5.inputs import sha,verify_pins


def read(path):
    return json.loads(Path(path).read_text(encoding='utf-8-sig'))


def write(name,value):
    (RUN/name).write_text(json.dumps(value,ensure_ascii=False,indent=2),encoding='utf8')


baseline=read(RUN/'baseline.json')
actual=read(RUN/'current_preview.json')
document=read(RUN/'new_document.json')
tests=read(RUN/'targeted_tests.json')
s5r1=read(RUN/'s5r1_targeted.json')
native=read(RUN/'native_word.json')
visual=read(RUN/'visual_validation.json')
assert tests['successful'] and tests['tests_run']==16
assert s5r1['successful'] and s5r1['tests_run']==4
assert native['status']=='PASS_NORMAL_READONLY_OPEN' and native['unchanged']
assert native['owned_instance'] and visual['status']=='PASS'
assert sha(document['output'])==document['sha256']==visual['output_sha256']
assert not tests['network_attempts'] and not s5r1['network_attempts']
verify_pins()

current={str(p):sha(p) for p in [*product_paths(),ROOT/'STAGE5_DEV.md',ROOT/'tests/test_stage5.py',
    ROOT/'tests/test_stage5_s5_r1.py',ROOT/'tests/test_stage5_sample_results.py']}
assert current==read(RUN/'product_before_tests.json')==read(RUN/'product_after_tests.json')
assert current==read(RUN/'s5r1_product_before.json')==read(RUN/'s5r1_product_after.json')
for path in product_paths():
    assert current[str(path)]==document['product_binding'][str(path)],('Runtime drift after generation',path)
for path in current:
    if Path(path).suffix=='.py':
        ast.parse(Path(path).read_text(encoding='utf8'),filename=path)

allowed={str(ROOT/p) for p in ('uzk_stage5/word.py','uzk_stage5/gui.py','STAGE5_DEV.md')}
protected=baseline['protected_before']
assert not any(Path(path).name.casefold() in ('взмк.xlsx','узк.xlsm') for path in protected)
protected_after={path:sha(path) for path in protected}
changes={path for path,h in protected.items() if protected_after[path]!=h}
assert changes==allowed,('Unexpected protected changes',changes^allowed)
sources={path:h for path,h in protected.items() if path not in baseline['product_before']}
assert all(protected_after[path]==h for path,h in sources.items())
old=baseline['existing_test_docx']
assert sha(old['path'])==old['sha256']
assert len(actual['preview']['marks'])==4 and actual['preview']['controlled_units']==33

write('product_before.json',baseline['product_before'])
write('product_after.json',read(RUN/'product_after_tests.json'))
write('product_current.json',current)
write('protected_after.json',protected_after)
write('source_safety.json',{'status':'PASS','protected_before_count':len(protected),
    'unchanged_protected_files':len(protected)-len(changes),'authorized_changed_paths':sorted(changes),
    'protected_source_and_old_evidence_count':len(sources),'all_source_hashes_unchanged':True,
    'old_test':old,'old_test_unchanged':True,'historical_acts_unchanged':10,
    'production_workbooks_opened':False,'production_hashing':'NOT_RUN_BY_POLICY',
    'stage1_4_changed':False,'registry_decrement':False,'passport_closed':False})

descriptions={
    'uzk_stage5/result_samples.py':'Added: exact ten-source pins, literal extraction, uniqueness proof, TEST-only gate.',
    'uzk_stage5/word.py':'Changed: two sample cells and visible TEST/unconfirmed title, source recheck.',
    'uzk_stage5/gui.py':'Changed: one banner string, results unconfirmed instead of results empty.',
    'STAGE5_DEV.md':'Changed: sample contract, future per-mark confirmation gate, current evidence instructions.',
    'tests/test_stage5_sample_results.py':'Added: 16 targeted checks; new validation fixtures/evidence only.'}
write('changed_files.json',[{'path':str(ROOT/p),'change':why,
    'before':baseline['product_before'].get(str(ROOT/p)),'after':current[str(ROOT/p)]}
    for p,why in descriptions.items()])

# A matching old source copy, when retained by prior evidence, gives a reviewable
# diff without recovering or modifying any previous product or evidence.
patches=[]
for target in sorted(allowed):
    expected=baseline['product_before'][target]
    copies=[Path(path) for path,h in protected.items() if h==expected and path!=target]
    if copies:
        original=copies[0]
        assert sha(original)==expected
        patches.extend(difflib.unified_diff(original.read_text(encoding='utf8').splitlines(True),
            Path(target).read_text(encoding='utf8').splitlines(True),fromfile='BEFORE/'+str(Path(target).relative_to(ROOT)),
            tofile='CURRENT/'+str(Path(target).relative_to(ROOT))))
if patches:
    (RUN/'TARGETED_CHANGE.diff').write_text(''.join(patches),encoding='utf8')

hash_lines=['# Product hashes BEFORE AFTER CURRENT','',
    'BEFORE is the accepted S5-R1 product. AFTER and CURRENT equal the final test-bound product.','',
    '| FILE | BEFORE | AFTER | CURRENT |','|---|---|---|---|']
for path,h in current.items():
    hash_lines.append('| '+str(Path(path).relative_to(ROOT))+' | '+baseline['product_before'].get(path,'ADDED')+' | '+h+' | '+h+' |')
(RUN/'PRODUCT_HASHES.md').write_text('\n'.join(hash_lines)+'\n',encoding='utf8')

history=(RUN/'HISTORICAL_TEXT.md').read_text(encoding='utf8').split('\n',2)[2]
file_lines='\n'.join('- `'+path+'`: '+reason for path,reason in descriptions.items())
report=f'''# Stage 5 targeted change — sample result columns

## 1. ACTUAL STATE

Accepted Stage 5 and S5-R1 independent rereview verified: PASS, S5-R1 CLOSED.
All 11 accepted product hashes matched before editing. The four previous
implementation/review manifests and their evidence files were verified unchanged.
Previous user preview E2E: USER_REPORTED_PASS. Current targeted Work review: pending.

## 2. HISTORICAL TEXT — ОБЕ ГРАФЫ

Only ten actual Garage Kazakhstan acts were examined for result text: 56 rows.
Exact raw text is «ДНО» / «годен», with original case retained. No expansion
of «ДНО» was invented. Source SHA, table and physical row are retained in
historical_rows.json and historical_sources.json; CSV contains all four requested columns.

{history}

## 3. PATTERN ASSESSMENT

One exact pair in all 56 rows / 10 acts. Blank results: 0. Different variants: 0.
Actual negative conclusions: 0. These cells describe the historical inspections;
they are not evidence of a new shipment's actual inspection or suitability.
Pattern accepted solely for SAMPLE_ONLY mode.

## 4. CHANGED FILES

{file_lines}

New evidence, helpers and isolated regression fixtures are confined to this run.
core.py, inputs.py, Stage 1–4, existing test suites and prior evidence are unchanged.

## 5. IMPLEMENTATION / STOP REASON

TEST/DRAFT only. Each generation reads exactly the ten accepted paths and verifies
their hashes, all 56 literal rows and the unique pair. Missing, changed, empty,
invalid or conflicting samples stop generation. Source hashes are checked again
before exclusive output creation. Other projects have no accepted sample and stop.

Both labels are visibly printed in the title:
ТЕСТ / ОБРАЗЕЦ ЗАПОЛНЕНИЯ
РЕЗУЛЬТАТЫ КОНТРОЛЯ НЕ ПОДТВЕРЖДЕНЫ

Official mode is rejected. Its future contract requires each mark's actual two
results, responsible executor, confirmation time and inspection evidence.
The existing quantity-check checkbox does not confirm results. Official issuance
and its confirmation workflow are not implemented.

## 6. NEW TEST DOCX PATH + SHA

Path: `{document['output']}`

SHA256: `{document['sha256']}`

The previous accepted TEST DOCX is unchanged: `{old['sha256']}`.

## 7. WORD VISUAL VALIDATION

PASS. Installed Microsoft Word {native.get('version')} / build {native.get('build')};
normal read-only open, no repair mode, owned process only, close without saving.
Two pages inspected completely; no clipping, no additional blank pages, no hyperlinks.
Exactly four data rows; both sample cells filled on every row; both labels visible.
The preserved second page contains the executor, empty signature line and qualification.
The user's already-open Word process and its document were left open and unchanged.

## 8. TARGETED TESTS

16/16 sample checks PASS; 4/4 selected S5-R1 cases PASS. No network attempts.
Independent python-docx extraction equals all 56 raw records from separate OOXML extraction.
Source drift, missing/blank/ambiguous examples and official mode stop.
Against the old TEST, only word/document.xml changes. After replacing the authorized
title and columns 7/8 with their old nodes, document.xml is byte-identical.
First six cells, all page/table/section settings, equipment, welders, executor and
signature content are unchanged.

| MARK | Q | L × S, mm | FREQUENCY / ANGLE |
|---|---:|---|---|
| Бп2-1 | 8 | 1280×5; 1024×8.4 | 5 MHz / 70° |
| Пс1-9 | 9 | 3600×4 | 5 MHz / 70° |
| Пс1-10 | 15 | 6000×4 | 5 MHz / 70° |
| Пс2-16 | 1 | 480×3 | 5 MHz / 70° |

Total: 4 marks / 33 units. Passport 9/2026, date 01.10.2026 unchanged.
S5-R1: empty intersection does not load profiles, generate output, or enable UI
generation; unrelated profile conflicts remain ignored; required registry conflicts stop.

## 9. SOURCE SAFETY

PASS. {len(sources)} protected source/old-evidence hashes unchanged.
{len(protected)-len(changes)} of {len(protected)} baseline files unchanged; the only
three changed baseline files are the authorized word.py, GUI banner and DEV guide.
Two new product/test files were added. All ten historical acts and both old TEST
locations, including the user's current DEV preview document, remain unchanged.
DEV workbooks, project mark catalog, Profiles.xlsx and all Stage 1–4 files unchanged.
Production workbooks were not opened or hashed. No decrement or passport closure.

## 10. EVIDENCE

PRODUCT_HASHES.md and product_before/after/current.json bind the accepted baseline
and final current files. Final product_before_tests = product_after_tests = CURRENT;
the four S5-R1 checks bind the same final bytes. Runtime hashes equal generation-time
hashes. Documentation/test-harness preparation changes are recorded separately.
Native Word JSON, both page PNGs, visual_validation.json, source_safety.json,
targeted/S5-R1 logs, access guards and exact source extractions are included.
Preparation attempts and resolved helper/test-harness errors are retained in
PREPARATION_NOTES.md and separate preparation evidence. No old evidence was edited.
REVIEW_MANIFEST.json is written last and includes SHA and size of every evidence file.

## 11. NOT_RUN

- Actual ultrasonic inspection and per-mark executor confirmation.
- Official issuance, registry decrement, passport closure, production generation.
- New independent targeted Work review and user E2E after this change.
- Full Stage 1–4/Stage 5 regression; only required targeted checks ran.
- LibreOffice renderer unavailable (soffice.exe missing); native Word render PASS.
- Stage 6, web research and Stage 3 import.

## 12. READY FOR TARGETED WORK REVIEW — YES

Implementation and required targeted/native validation complete; current change
is ready for independent Work review. Stop after this report. Stage 6 not started.
'''
(RUN/'REPORT.md').write_text(report,encoding='utf8')

# The manifest is the last write. Verification below reads only.
manifest_path=RUN/'REVIEW_MANIFEST.json'
assert not manifest_path.exists(),'Never overwrite a finalized review manifest'
evidence=[{'path':str(p.relative_to(RUN)),'sha256':sha(p),'size':p.stat().st_size}
          for p in sorted(RUN.rglob('*')) if p.is_file()]
manifest={'status':'PASS_TARGETED_CHANGE','ready_for_targeted_work_review':True,
    'created_utc':datetime.now(timezone.utc).isoformat(),'baseline_review':'S5-R1 CLOSED / PASS',
    'current_work_review':'NOT_RUN','mode':'SAMPLE_ONLY','results_confirmed':False,
    'product_before':baseline['product_before'],'product_after':current,'product_current':current,
    'test_binding_before':read(RUN/'product_before_tests.json'),
    'test_binding_after':read(RUN/'product_after_tests.json'),
    'tests':{'sample':16,'s5r1':4,'successful':True},'output':document['output'],
    'output_sha256':document['sha256'],'native_word_validation':'PASS',
    'source_safety':'PASS','source_hashes_unchanged':len(sources),
    'review_manifest_written_last':True,'evidence':evidence}
write('REVIEW_MANIFEST.json',manifest)
for item in evidence:
    path=RUN/item['path']
    assert path.stat().st_size==item['size'] and sha(path)==item['sha256'],path
assert all(p.stat().st_mtime_ns<=manifest_path.stat().st_mtime_ns for p in RUN.rglob('*') if p.is_file())
assert all(sha(path)==expected for path,expected in current.items())
assert all(sha(path)==expected for path,expected in sources.items())
print(json.dumps({'status':manifest['status'],'tests':'20/20 PASS','evidence_files':len(evidence),
    'source_hashes_unchanged':len(sources),'old_test_unchanged':True,'manifest':str(manifest_path),
    'manifest_sha256':sha(manifest_path),'output_sha256':document['sha256'],
    'review_manifest_last':True,'ready':True},ensure_ascii=False,indent=2))
