"""Record final independent review and preservation gate without product writes."""
import json, hashlib
from pathlib import Path
OUT=Path(__file__).resolve().parent;ROOT=OUT.parents[1]
def read(p):return json.loads(Path(p).read_text(encoding='utf-8-sig'))
def sha(p):return hashlib.sha256(Path(p).read_bytes()).hexdigest()
def save(n,d):(OUT/n).write_text(json.dumps(d,ensure_ascii=False,indent=2),encoding='utf-8')
before=read(OUT/'reviewed_hashes_before.json');current={p:sha(p) for p in before}
assert before==current
assert read(OUT/'reviewed_hashes_after_inspect.json')==before
assert read(OUT/'reviewed_hashes_after_fixtures.json')==before
impl=ROOT/'validation/stage4_implementation_20261006_223330'
manifest=read(impl/'STAGE4_REVIEW_MANIFEST.json')
fixture=read(OUT/'fixture_summary.json');assert fixture['reviewed_originals_unchanged']
assert len(read(OUT/'valid_external_race.json'))==3
save('reviewed_hashes_final.json',current)
findings=[
 {'id':'M1','severity':'MAJOR','file':'profiles_stage4/model.py','function':'Catalog.__init__/lookup',
  'lines':[162,166,175,187,196,222],
  'problem':'Every disagreement, including area/mass-only and R/r-only, blanks all canonical geometry and blocks lookup.',
  'actual_impact':'7 current valid main geometries are blocked solely by area/mass differences. Same-ND radius-only fixture also blocks all working geometry.',
  'evidence':['conflicts_independent.json','policy_probes.json'],
  'minimal_fix':'Separate blocking main-geometry/family differences from optional/reference differences. Retain agreed main geometry; keep disputed R/r and quality unresolved and traceable without automatic source selection.',
  'targeted_validation':'All 15 current conflicts; 7 reference-only rows usable; 8 main-geometry rows still blocked; same-ND R-only/r-only fixtures; preferred same-ND geometry conflict stays blocked; existing golden tests.'},
 {'id':'M2','severity':'MAJOR','file':'profiles_stage4/storage.py','function':'build',
  'lines':[177,178,193,205,209,214,215],
  'problem':'Integrity is checked before rendering only. Valid external edits arriving during rendering are silently overwritten; their backup no longer matches its build manifest.',
  'actual_impact':'Independently reproduced for workbook, catalog JSON and authoritative user JSON in isolated fixtures. build returned success and discarded valid external edits from active state.',
  'evidence':['fixture_results.json','valid_external_race.json'],
  'minimal_fix':'Capture and recheck the complete previous generation at commit, protect replacement against concurrent external changes, and STOP without overwriting changed bytes. Ensure backup is a verified complete generation.',
  'targeted_validation':'Valid Excel/catalog/user edits during rendering and commit must STOP and retain edits; legitimate service writes/rebuild must remain allowed; lock, rollback, interruption and state-survival tests.'}
]
notes={'review_only_harness_repairs':['Renamed inspect.py to avoid shadowing Python inspect.',
 'Handled display numeric non-breaking spaces and datetime evidence serialization.',
 'Corrected a traversal test to actually escape validation; resumed remaining tests.',
 'Corrected workbook notes column T2 for valid external edit reproduction.'],
 'product_files_changed':False,'original_implementation_suite_executed':False,
 'original_suite_reason':'Contains real DEV build and writes to original implementation evidence.',
 'fault_transaction_renderer':'Byte-identical verified workbook copied into staged output after exact payload equality check; real renderer exercised in state rebuild tests.'}
save('harness_notes.json',notes)
result={'verdict':'FAIL','severity_counts':{'BLOCKER':0,'MAJOR':2,'MINOR':0},
 'artifact':manifest['workbook'],'workbook_sha256':manifest['workbook_sha256'],
 'runtime_product_hashes':{p:h for p,h in manifest['tested_product_hashes'].items() if not p.startswith('validation')},
 'evidence_tool_hashes':{p:h for p,h in manifest['tested_product_hashes'].items() if p.startswith('validation')},
 'sidecars':{'profiles.user.json':manifest['user_state_sha256'],'profiles.catalog.json':manifest['catalogue_sha256'],'profiles.build.json':manifest['commit_sha256']},
 'DEV_registry_sha256':sha(ROOT/'DEV/ВЗМК_DEV.xlsx'),'originals_checked':len(before),'originals_unchanged':True,
 'source_evidence_hashes':{rel:sha(ROOT/'validation'/rel/'AUDIT_MANIFEST.json') for rel in ['stage4_km_kmd_audit_20261006_204347','stage4_ptoflow_audit_20261006_215239']},
 'source_verification':read(OUT/'source_verification.json'),'counts':read(OUT/'counts.json'),
 'findings':findings,'fixture_summary':fixture,'valid_external_race':read(OUT/'valid_external_race.json'),
 'gui':read(OUT/'gui_independent.json'),'evidence_integrity':read(OUT/'evidence_independent.json'),
 'production_access':0,'network_attempts':0,'READY_FOR_USER_E2E':False,
 'NOT_RUN':['User E2E','Native Excel visual UAT','full Stage 1-3 workflows','Stage 3 Apply','production access','Stage 5','passport/VIK/UZK DOCX generation','live source research','abrupt process/power termination'],
 'harness_notes':notes}
result['review_evidence_files']={str(p.relative_to(OUT)):{'sha256':sha(p),'size':p.stat().st_size} for p in sorted(OUT.glob('*')) if p.is_file() and p.name!='INDEPENDENT_REVIEW.json'}
save('INDEPENDENT_REVIEW.json',result)
print(json.dumps({'verdict':result['verdict'],'findings':result['severity_counts'],'preserved':len(before),'workbook_sha256':result['workbook_sha256']},ensure_ascii=False))
