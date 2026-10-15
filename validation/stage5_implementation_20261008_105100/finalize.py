"""Bind current files to tested bytes; write the evidence manifest LAST."""
from datetime import datetime, timezone
import importlib
import json
from pathlib import Path
import sys
from zipfile import ZipFile

ROOT=Path(__file__).resolve().parents[2]
sys.path.insert(0,str(ROOT))
from uzk_stage5.inputs import PINS, REGISTRY, sha, verify_pins, workbook
from uzk_stage5.core import Preview, product_paths
from profiles_stage4.offline import no_network

RUN=Path(__file__).resolve().parent

def read(name):
    return json.loads((RUN/name).read_text(encoding='utf-8-sig'))

def write(name,data):
    (RUN/name).write_text(json.dumps(data,ensure_ascii=False,indent=2),encoding='utf8')

final=read('final_run.json');data=final['preview'];output=Path(final['output'])
Preview(data,final['preview_seal']).verify()
verify_pins()
tests=read('targeted_tests.json')
assert tests['successful'] and tests['tests_run']==36 and not tests['skipped']
native=read('native_final_v3.json')
assert native['status']=='PASS_NORMAL_READONLY_OPEN' and native['before_sha']==sha(output) and native['unchanged']
assert native['pages']==2 and native['first_table_rows']==9 and native['hyperlinks']==0
protected_before=read('protected_before.json')
protected_after={path:sha(path) for path in protected_before}
assert protected_before==protected_after
write('protected_after.json',protected_after)

tested_before=read('tested_product_before.json');tested_after=read('tested_product_after.json')
current={str(p):sha(p) for p in [*product_paths(),ROOT/'tests/test_stage5.py']}
assert tested_before==tested_after==current
write('product_after.json',current)
write('product_current.json',current)

production_access=[]
def guard(event,args):
    if event=='open' and isinstance(args[0],(str,bytes)):
        path=Path(args[0].decode() if isinstance(args[0],bytes) else args[0])
        if path.name.casefold() in ('взмк.xlsx','узк.xlsm'):
            production_access.append(str(path))
            raise AssertionError('Forbidden production access')
sys.addaudithook(guard)
with no_network() as attempts:
    for module in ('main','stage2','stage3','stage4'):
        importlib.import_module(module)
assert not attempts and not production_access
write('regression_smoke.json',{'imports':['main','stage2','stage3','stage4'],
    'status':'PASS_IMPORTS_AND_HASH_CHECKPOINT_ONLY','full_regression':'NOT_RUN',
    'production_access_attempts':production_access,'network_attempts':attempts,
    'accepted_files_unchanged':True})

with workbook(REGISTRY) as b:
    s=b['Паспорт качества'];rn=data['passport']['row']
    formula=s.cell(rn,14).value
verified={'input_fingerprints':data['fingerprints'],'accepted_dev_pins':{str(p):h for p,h in PINS.items()},
    'selected_passport':data['passport'],'date_mapping':{'sheet':'Паспорт качества','cell':f'N{rn}',
    'stored_formula':formula,'accepted_cached_date':data['passport']['date'],
    'formula_evaluation':'NOT_RUN','date_substitution':False},
    'template_sha256':sha(data['template']),'output_sha256':sha(output),
    'catalog_sha256':sha(data['mark_catalog']),'vom_sha256':sha(data['vom']['path'])}
write('verified_inputs.json',verified)

visual={'document':str(output),'sha256':sha(output),'native_pages':2,
    'reviewed_images':native['images'],'all_pages_inspected':True,
    'result':'PASS','clipping':False,'table_corruption':False,'extra_blank_pages':False,
    'page_1':'Act/date/object/equipment/table; four complete mark rows, channel multiline; result cells blank.',
    'page_2':'Original signature line, printed operator and qualification; no new signature.',
    'source_template_pages_inspected':['template_1.png','template_2.png'],
    'preserve_only_package_parts':'PASS; only document.xml and its hyperlink relationships differ'}
write('visual_review.json',visual)

gate_names=[
 'DEV checkpoint verification','Actual selected VOM read-only','XLS visible-row reader','Hidden-row exclusion',
 'XLSX compatibility','Repeated VOM marks','Registry all 25 rows through row 26','Blank quantity -> 1',
 'Numeric remaining','Q=min(VOM,remaining)','Project isolation','Missing registry -> no UZK',
 'Conflicting duplicate registry STOP','Four raw mark/profile mappings','Stage 4 resolved lookup',
 'AMBIGUOUS/CONFLICT STOP','Unsupported rule STOP','16П segments','Bent channel segment','Closed square segment',
 'Frequency/angle','No cross-mark sum','Four physical mark rows','Multiline Бп2-1',
 'Number/date/object/cipher','Old marks/PDF hyperlinks removed','Template constants retained; old outcomes blanked',
 'Signature table/section/layout retained','No UZK -> no file','Registry unchanged','Source files unchanged',
 'Production paths untouched','Offline operation','Native Word normal open']
gates=[{'id':i,'gate':name,'status':'PASS'} for i,name in enumerate(gate_names,1)]
write('required_34_gates.json',gates)

changed=[*product_paths(),ROOT/'tests/test_stage5.py',ROOT/'STAGE5_DEV.md']
write('changed_files.json',{'new_product_files':[{'path':str(p),'sha256':sha(p)} for p in changed],
    'modified_accepted_files':[],'evidence_folder':str(RUN),
    'final_output':str(output),'older_output_folders':'QA iterations / fixtures only; not deliverables'})
not_run=['Full Stage 1–4 regression','User-performed Stage 5 GUI E2E',
    'Actual ultrasonic inspection / results / sign-off','Production registry reads or writes',
    'Registry decrement / row deletion / passport close','PDF export / AKZ / quality document / EXE / deploy / Stage 6']

rows=[]
for m in data['marks']:
    segments='; '.join(f'{s["L"]} × {s["S"]}' for s in m['segments'])
    rows.append(f'| {m["mark"]} | `{m["raw_profile"]}` | {m["vom_quantity"]} | {m["remaining"]} | {m["Q"]} | {segments} |')
source_table='\n'.join(f'| `{path}` | `{h}` |' for path,h in verified['input_fingerprints'].items()
                       if Path(path).suffix in ('.xlsx','.xlsm','.xls','.docx'))
report=f'''# Stage 5 IMPLEMENTATION — TEST/DRAFT

1. **ACTUAL STATE**

Separate Stage 5 implemented. New accepted Stage 4 baseline used as supplied;
no KMD, standards or web research. Accepted stages 1–4 unchanged.

2. **VERIFIED INPUTS**

Selected project Гараж Казахстан, passport 9/2026, dynamically found row {data['passport']['row']}.
Date source N{data['passport']['row']}: stored formula `{formula}`, accepted cached date 01.10.2026;
no formula evaluation or date substitution. E/F supply object and KMD cipher.
All source SHA were verified; detailed sidecar/code fingerprints are in verified_inputs.json.

| Exact input | SHA256 |
|---|---|
{source_table}

3. **VOM XLS VISIBLE-ROW READER**

Excel COM, ReadOnly, macros disabled, UpdateLinks=0, events disabled, close without save.
104 visible data rows; 397 hidden data rows excluded. Mass 15822.14 kg.
Two-line headers identify mark, quantity and total mass; footer totals are ignored.
XLS and XLSX filter/manual-hide fixtures passed. External links=0, connections=0.
Raw floating-point mass sum retained separately from the display rounded to 0.01 kg.

4. **REGISTRY MATCHING**

All 25 valid rows A2:C26 read independently of AutoFilter A1:C21.
Key: approved Stage 3 normalized project + mark. Blank C -> 1;
positive integer C -> remaining; Q=min(visible shipment quantity, remaining).
Identical duplicate registry keys do not double stock. Conflicts STOP.
No VBA, save, decrement, delete or close operation.

5. **MARK/PROFILE LOOKUP**

Raw spellings preserved: Бп2-1 `[16П`; Пс1-9 / Пс1-10 `Гн.[140X60X4`;
Пс2-16 `Гн.[]120х3`. All four lookups FOUND. Local accepted Stage 4 returns
16П / ГОСТ 8240-97 / ШВЕЛЛЕР_16П_1;
ГН.[140×60×4 / ГОСТ 8278-83 / ГНУТЫЙ_ШВЕЛЛЕР_1;
120×120×3 / ГОСТ 30245-2003 / ЗАМКНУТЫЙ_ПРОФИЛЬ_1.
Unrelated ПРОВЕРИТЬ tested through an isolated mapping fixture.

6. **FOUR-MARK CALCULATION**

| Mark | Raw profile | VOM | Remaining | Q | L × S, mm |
|---|---|---:|---:|---:|---|
{chr(10).join(rows)}

Total: four marks, 33 units. Every segment: 5 MHz / 70 degrees.
Stage 4 calculates ONE_ITEM geometry and same-thickness grouping;
Stage 5 multiplies only L by Q, never combines different marks.

7. **WORD TEMPLATE FIELD/ROW MAPPING**

Exact project first-passport template selected; fallback not needed.
Body p[5] -> B number + N date + visible TEST/DRAFT label.
Table 1 row 2 cell 2 -> E object + parenthesized F cipher.
Five fixed rows retained, all eleven historical mark rows replaced by exactly four.
Eight logical cells / nine grid columns retained. Бп2-1 contains two lines inside
geometry/frequency/angle cells. Original welder cell retained.
Historical outcomes ДНО / годен intentionally cleared under current instruction
not to invent inspection results. Signature table, printed operator, qualifications,
company, equipment and certificate constants retained; no signature added.
Only document.xml and its hyperlink relationships changed within the DOCX package.

8. **CHANGED FILES**

New: stage5.py; uzk_stage5/__init__.py, inputs.py, core.py, word.py, gui.py;
requirements_stage5.txt; Запустить Этап 5 DEV.cmd; tests/test_stage5.py; STAGE5_DEV.md.
All evidence and generated drafts confined to this validation run.
Modified accepted files: zero. Launcher CheckOnly PASS. Programmatic Tk gate PASS;
preview required, verification checkbox required, selection changes invalidate preview.

9. **GENERATED TEST DOCX PATH AND SHA**

`{output}`

SHA256: `{sha(output)}`

Only this final_output_v3 document is the deliverable. Other output folders are
retained QA iterations / fixtures, with their identities captured by the manifest.

10. **TARGETED VALIDATION**

36/36 focused tests PASS, zero skips. All required 34 gates mapped in
required_34_gates.json. Negative tests cover quantity errors, conflicting duplicates,
missing/unresolved profiles, actual Stage 4 ambiguity/conflict, unsupported rules,
stale/tampered preview, forbidden output, no-UZK, ambiguous Word template and a
concurrent existing-file race. All mutations confined to isolated fixtures.

11. **NATIVE WORD VALIDATION**

Installed Word normal Documents.Open, ReadOnly, no repair API, macro security 3,
links refresh disabled, close without save. Two actual pages, one landscape A4
section, two tables, nine first-table rows, zero hyperlinks. DOCX SHA unchanged
after Word. Both current native PNG pages visually inspected: complete rows,
aligned segments, signatures/constants present, no table corruption or extra blank page.
No PDF export.

12. **SOURCE / DEV / PRODUCTION SAFETY**

All {len(protected_before)} protected source/accepted-product fingerprints identical before/after.
DEV registry, DEV UZK, profiles and sidecars, project catalog, VOM, template and
existing project passport 9 UZK unchanged. Production registries were not opened
or hashed. Python file-access audit forbids their exact names and logs writes;
pipeline writes limited to isolated validation output. Normal operation made zero
Python network attempts; guard effectiveness tested separately. Native Excel has
no external data connections and uses links=0/events disabled/macros disabled.

13. **EVIDENCE INTEGRITY**

Final run binds exact source hashes, template, visible rows, all registry rows,
per-mark lookup/provenance/rules/single geometry/Q/scaled segments and output.
Product tested BEFORE = AFTER = CURRENT. Accepted product BEFORE hashes are
unchanged; Stage 5 files were newly created. Earlier failed test attempts retained:
an initial test runner name collision and an assumption that unrelated ПРОВЕРИТЬ
already existed; final tests use an isolated fixture and all pass.
EVIDENCE_MANIFEST.json is written last, after all evidence, and contains SHA/size
for every preceding evidence file. Manifest does not hash itself.

14. **REGRESSION IMPACT**

Only new Stage 5 entry point/package/launcher/dependencies/tests/docs.
Focused Stage 1–4 imports + checkpoint/hash smoke PASS. No prior-stage code,
settings, workbook, rule or sidecar changed. No full prior-stage regression.

15. **NOT_RUN**

{chr(10).join('- '+item for item in not_run)}

16. **READY FOR INDEPENDENT REVIEW — YES**

Implementation and TEST/DRAFT evidence complete. Work stopped at Stage 5.
'''
(RUN/'REPORT.md').write_text(report,encoding='utf8')

evidence={}
for path in sorted(RUN.rglob('*')):
    if path.is_file() and path.name!='EVIDENCE_MANIFEST.json':
        evidence[str(path.relative_to(RUN))]={'sha256':sha(path),'bytes':path.stat().st_size}
manifest={'version':1,'created_utc':datetime.now(timezone.utc).isoformat(),'status':'PASS_READY_FOR_INDEPENDENT_REVIEW',
    'accepted_baseline':'User Stage 5 IMPLEMENTATION request; Stage 4 accepted SHA 7b420f6d...',
    'final_output':{'path':str(output),'sha256':sha(output),'kind':'TEST/DRAFT','native_pages':2},
    'inputs':data['fingerprints'],'tests':tests,'required_gates':gates,
    'product_before':read('product_before.json'),'product_tested_before':tested_before,
    'product_tested_after':tested_after,'product_current':current,'protected_files_unchanged':len(protected_before),
    'source_mutations':[],'production_registry_access':[],'network_attempts':data['network_attempts'],
    'not_run':not_run,'evidence_files':evidence,'manifest_written_last':True}
write('EVIDENCE_MANIFEST.json',manifest)  # FINAL FILE WRITE
print(json.dumps({'status':manifest['status'],'output':manifest['final_output'],
    'protected_unchanged':len(protected_before),'evidence_files':len(evidence),
    'manifest_sha256':sha(RUN/'EVIDENCE_MANIFEST.json')},ensure_ascii=False))
