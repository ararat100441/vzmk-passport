import json,hashlib,re
from pathlib import Path
from datetime import datetime,timezone
OUT=Path(__file__).resolve().parent;APP=OUT.parent.parent
def read(p):return json.loads(Path(p).read_text(encoding='utf-8-sig'))
def save(n,x):(OUT/n).write_text(json.dumps(x,ensure_ascii=False,indent=2),encoding='utf-8')
def sha(p):return hashlib.sha256(Path(p).read_bytes()).hexdigest()
def table(head,rows):
 def cell(v):return str(v).replace('|','/').replace('\n',' / ')
 return '\n'.join(['| '+' | '.join(head)+' |','| '+' | '.join('---' for _ in head)+' |']+['| '+' | '.join(cell(v) for v in r)+' |' for r in rows])
reports=read(OUT/'project_uzk_documents.json');voms=read(OUT/'historical_vom_quantities.json')
before=read(OUT/'preservation_before.json')
for v in voms:before[v['path']]=v['before_sha']
dev=str(APP/'DEV'/'ВЗМК_DEV.xlsx');before[dev]='8e56f68518c70cade9efb09084dd67cd5d72e07e3fdc813b8515030e9cb98e80'
after={p:sha(p) for p in before};assert before==after,'STOP input drift'
save('preservation_before.json',before);save('preservation_after.json',after)
catalog=read(OUT/'working_mark_catalog.json')
current_marks={'Бп2-1':'[16П','Пс1-9':'Гн.[140X60X4','Пс1-10':'Гн.[140X60X4','Пс2-16':'Гн.[]120х3'}
current_catalog=[r for r in catalog['rows'] if r['values'][0] in current_marks]
assert len(current_catalog)==4 and all(r['values'][1]==current_marks[r['values'][0]] for r in current_catalog)
save('current_catalog_proof.json',current_catalog)
targets=[]
for d in reports:
 for row in d['rows']:
  profile=(row.get('kmd') or {}).get('raw_profile','')
  if profile in ['[16П','Гн.[140X60X4','Гн.[200X100X6','Гн.[]120х3']:
   targets.append({'act':d['act'],'date':d['date'],'docx':d['path'],'table_row':row['table_row'],'mark':row['mark'],'profile':profile,'kmd_page':row['kmd']['page'],'detail':row['kmd']['detail'],'geometry':row['geometry_raw'],'frequency':row['frequency_raw'],'angle':row['angle_raw'],'links':row['hyperlinks']})
save('target_historical_examples.json',targets)
legacy=[]
for d in reports:
 for row in d['rows']:
  segs=re.findall(r'L\s*=?\s*([\d.,]+)\s*[хХxX×]\s*([\d.,]+)',row['geometry_raw'])
  f=row['frequency_raw'].splitlines();a=row['angle_raw'].splitlines()
  cardinality=len(segs)!=len(f) or len(segs)!=len(a)
  checks=[]
  for i,(length,s) in enumerate(segs):
   try:
    if s.endswith(('.',',')):raise ValueError('malformed thickness')
    s=float(s.replace(',','.'));expected=[5,70] if s<12 else [2.5,65]
    actual=[float(f[i].replace(',','.')),float(a[i].replace(',','.'))] if i<len(f) and i<len(a) else None
    checks.append({'S':s,'expected':expected,'actual':actual,'status':'MATCH' if actual==expected else 'LEGACY_MISMATCH'})
   except (ValueError,IndexError):checks.append({'status':'UNPARSEABLE_LEGACY_TEXT'})
  legacy.append({'act':d['act'],'mark':row['mark'],'segments':checks,'cardinality_mismatch':cardinality})
save('legacy_frequency_angle.json',legacy)
rules=[{'type':'Прокатный швеллер 16П','examples':3,'independent_marks':3,'quantity_resolved_examples':3,'candidate':'For verified 16П only: L1=h=160,S1=s=5; L2=2b=128,S2=t=8.4; multiply L by Q','status':'RULE_PROVEN','confidence':'HIGH within exact verified size; no extrapolation to other channels'},
 {'type':'Гнутый швеллер','examples':5,'independent_marks':5,'quantity_resolved_examples':3,'candidate':'L_single=2(h+b), S=t; 140×60×4 gives400×4;200×100×6 gives600×6','status':'CONTRADICTORY_EVIDENCE','confidence':'Two exact140 examples agree, but200 example Пп1-13 needs Q13 while corresponding VOM qty11'},
 {'type':'Квадратный замкнутый120×120×3','examples':10,'independent_marks':7,'quantity_resolved_examples':1,'candidate':'L_single=4×120=480,S=3','status':'INSUFFICIENT_EVIDENCE','confidence':'Only act9 Пс2-16 has independently established Q1; other historical Q unavailable'}]
save('rule_assessment.json',rules)
preview={'mark':'Бп2-1','Q':8,'rule_status':'RULE_PROVEN','scope':'EXACT_16П_ONLY','rows':[{'L_single':160,'S':5,'L_total':1280,'frequency_mhz':5,'angle_degrees':70},{'L_single':128,'S':8.4,'L_total':1024,'frequency_mhz':5,'angle_degrees':70}],'other_three_marks':'NOT_RUN_RULE_NOT_PROVEN','document_created':False}
save('current_four_mark_preview.json',preview)
summary={'status':'STOP_SOURCE_CONTRADICTION_AND_UNPROVEN_HISTORICAL_Q','ready_to_implement':False,'working_catalog_sha256':catalog['sha256'],'project_documents':len(reports),'project_document_mark_rows':sum(len(d['rows']) for d in reports),'targeted_examples':len(targets),'rule_assessment':rules,'protected_files':len(before),'protected_changes':[],'other_projects_read':0,'prior_inventory_usage':'HEADER_ONLY_FORMAT_CHECK; NO_OTHER_PROJECT_DOCUMENTS_OPENED','production_VZMK_read':False,'VBA_execution':False,'generation':False,'registry_mutation':False,'source_mutations':False,'preview_marks':1}
save('audit_summary.json',summary)
docs_table=table(['ACT','DATE','MARK ROWS','EXACT FILE','SHA256'],[[d['act'],d['date'],len(d['rows']),d['path'],d['sha256']] for d in reports])
evidence_table=table(['ACT','MARK','KMD PAGE / DETAIL','RAW PROFILE','L/S DOCUMENT'],[[e['act'],e['mark'],str(e['kmd_page'])+' / '+e['detail'],e['profile'],e['geometry']] for e in targets])
rule_table=table(['TYPE','EVIDENCE COUNT','RULE','STATUS','CONFIDENCE'],[[r['type'],str(r['examples'])+' rows / '+str(r['independent_marks'])+' marks / '+str(r['quantity_resolved_examples'])+' Q proven',r['candidate'],r['status'],r['confidence']] for r in rules])
report=f'''# Stage 5 SOURCE PROOF — 07.10.2026, Europe/Samara

## 1. ACTUAL STATE

Read-only audit. Working project mark catalogue now exists; previous absence is superseded. Exact path: {catalog['path']}. Current SHA256 BEFORE=AFTER: {catalog['sha256']}. Current four links confirmed on sheet Марки: Бп2-1 row40=[16П; Пс1-9 row78=Гн.[140X60X4; Пс1-10 row79=same; Пс2-16 row98=Гн.[]120х3. Raw spelling preserved. Accepted profiles SHA39f14b040d3a97365f9e44f5fdc01c87e1f2ce068d5d4f26458a91602759ca91, sidecars and Stage4 code unchanged. DEV УЗК copy and checkpoint match prior audit. {len(before)} protected sources BEFORE=AFTER.

Current Q is user-confirmed prior evidence: Пс2-16=1,Пс1-10=15,Пс1-9=9,Бп2-1=8. Registry not recalculated. No production ВЗМК access.

## 2. PROJECT UZK DOCUMENTS FOUND

All10 project УЗК DOCX inspected at OOXML level. Full row values, frequency/angle, dates and hyperlink targets: project_uzk_documents.json. Only first act has mark hyperlinks:11 targets under file:///C:/Users/arara/Модель/Гараж%20Казахстан/0_номер_заказа/3_КМД/PDF/Сборки/. These outside-project targets were not followed; corresponding marks were linked to adjacent in-project WAA0007144-AC2-КМД.pdf by specification. Later9 acts have no mark hyperlinks.

{docs_table}

All targeted rows and exact source paths are also in target_historical_examples.json:

{evidence_table}

## 3. HISTORICAL EVIDENCE — 16П

Act8,27.09.2026: Бп2-6 row7 and Бп2-13 row8 each160×5 +128×8.4. KMD pages66/73,details286/293 as recorded in project_kmd_mark_evidence.json (exact detail numbers below are authoritative in JSON). Each specification raw[16П; metal selection ГОСТ8240-97. Corresponding filtered VOM8 matches marks at rows41/48 with visible D=1 each; positive required count plus only one shipment item establishes Q1 without guessing historical remaining. Full visible VOM mass16541.88 matches DEV passport8. Other short VOM.xls/xlsx in same folder contain only Бп1-1…Бп1-32 and cannot establish these examples.

Act9,01.10.2026: Бп2-1 row9:1280×5 +1024×8.4; current Q8 independently confirmed. KMD page61/detail303/raw[16П/ГОСТ8240-97. Divide proven Q8→160×5 +128×8.4. Three distinct marks agree.

Accepted Stage4 geometry16П: h160,b64,s5,t8.4,r5,R8.5. Observed first L=h, without subtracting2t (h−2t=143.2 does not match160). Second L=2b=128, two flange lengths combined at equal t. No radius correction observed. RULE_PROVEN only within exact16П, not a general rule for every channel size.

## 4. HISTORICAL EVIDENCE — Гн.[140X60X4

Act9: Пс1-10 row7,6000×4,Q15→400×4; KMD page414/detail527. Пс1-9 row8,3600×4,Q9→400×4; KMD page413/detail536. Both rawГн.[140X60X4,ГОСТ8278-83; h140,b60,t4. Candidate400=2(h+b), rather than h+2b260. Two independently quantity-resolved exact140 examples.

Same-family check act8 uses rawГн.[200X100X6,ГОСТ8278-83: Пп1-1 page378,600×6, visible VOM D316=1→single600×6, matching candidate2(200+100). But Пп1-13 page390 has7800×6 while visible VOM D328=11. Candidate600 would require13 controlled items, exceeding11 shipment items. This cannot be reconciled by selecting a smaller historical registry remaining. Пп1-5 page382 has15000×6 and visible VOM D320=59; historical Q unavailable, so Q25 must not be inferred from L.

STOP: contradictory historical source chain for proposed bent-channel rule. No source corrected and no replacement Q guessed. Need authoritative explanation of act8/VOM8 discrepancy. Literal exact140 evidence alone remains only2 marks.

## 5. HISTORICAL EVIDENCE — Гн.[]120х3

Act7,26.09.2026: Пс2-3=480×3;Пс2-2=1440×3;Пс2-16=480×3;Пс2-1=2880×3;Пс2-18=960×3;Пс2-11=960×3;Пс2-17=480×3. Act8 repeats Пс2-3/Пс2-17 at480×3. Act9 repeats Пс2-16 at480×3. Ten rows,7 distinct marks, all rawГн.[]120х3 linked by assembly specification.

Only act9 Пс2-16 has proven current Q1, therefore proven single480×3. KMD page458/detail718 and Детали.pdf page531/sheet4.531 establish square closed section h=b120,t3,ГОСТ30245-2003. The main member is718; end plates6/4mm are separately listed. Radii are not specified and were not guessed.

VOM7 quantities: Пс2-1=8,Пс2-2=7,Пс2-3=8,Пс2-11=6,Пс2-16=7,Пс2-17=4,Пс2-18=5. VOM8 Пс2-3=7,Пс2-17=4. These shipment counts do not prove historical controlled Q. The corresponding historical registry snapshots/explicit act counts are unavailable. Inferred divisors1/2/3/6 from480 multiples are prohibited. Candidate4×120 is consistent with one confirmed single-item example, but INSUFFICIENT_EVIDENCE. It does not prove whether the recorded length represents four individual sides, a perimeter, multiple welds or another inspection convention.

## 6. RULE ASSESSMENT

{rule_table}

Geometry source chain and rule meaning are confined to historical document-writing behavior, not a new technical inspection standard or actual inspection approval.

## 7. SAME-THICKNESS ASSESSMENT

16П:128×8.4=2×64 at the same thickness in one document line, consistent across3 marks with known Q. No different marks merged. Bent channels show one line at t4 or6, but decomposition into physical sections is not proven and candidate has conflict. Square has one S3 line, but historical Q gap prevents general automatic aggregation rule proof. Existing accepted same-S aggregation policy remains unchanged; do not claim physical segments proved from one combined number.

## 8. LEGACY FREQUENCY/ANGLE ASSESSMENT

Approved S<12→5/70;S>=12→2.5/65 retained. Act7 Пс2-17,S3 has2.5/65 (twice):LEGACY_MISMATCH and segment-count mismatch. Act9 Пс1-9,S4 has2.5/65:LEGACY_MISMATCH. Act9 Бп2-1,S5 and8.4 both2.5/65:LEGACY_MISMATCH. Act9 Пс2-16,S3 has correct5/70 duplicated despite one geometry line:values MATCH, cardinality mismatch. Act8 targeted values MATCH. Full per-row checks in legacy_frequency_angle.json. Malformed legacy thickness text is marked UNPARSEABLE, not silently corrected. No historical file edited.

## 9. STAGE 4 LOOKUP GAP DIAGNOSIS

Current read-only lookup: [16П→NOT_FOUND;16П+ГОСТ8240-97→FOUND;Гн.[140X60X4+ГОСТ8278-83→NOT_FOUND;Гн.[]120х3→NOT_FOUND;120×120×3 with and without ГОСТ30245-2003→NOT_FOUND. Only16П has resolved geometry; its current uzk_rule is НЕ НАСТРОЕНО. No rule was registered.

A Future exact/safe alias[16П→16П is needed for lookup while preserving raw display. Current normalize_profile does not strip[. Do not use indiscriminate punctuation removal.
B No ГОСТ8278 rows exist in accepted raw catalogue. Bent140×60×4 needs a source-grounded geometry entry and distinct rolled/bent identity; do not map it to rectangular closed140×60×4. KMD proves h140,b60,t4, not radii. No missing dimension extrapolated.
C Square family exists, but key120×120×3 does not exist in either accepted raw records or canonical rows. At h=b120, thicknesses are3.5,4,4.5,5,5.5,6,6.5,7;3 is absent. Thus this is a missing-size data gap in addition to shorthand normalization, not mere absence of square family. Neighbor120×120×3.5 is not a substitute. This audit does not establish whether the complete ГОСТ permits120×120×3; only project KMD declaration and current offline catalogue contents are established. Future verified entry and explicit shorthand alias are separate from rule proof.

## 10. CURRENT FOUR-MARK PREVIEW

Only RULE_PROVEN applied: Бп2-1,Q8: single160×5→total1280×5,5MHz/70°;single128×8.4→total1024×8.4,5MHz/70°. S not multiplied. This differs from historical act9 frequency/angle and uses approved rule.

Пс1-9,Q9;Пс1-10,Q15;Пс2-16,Q1:preview NOT_RUN because respective rule notRULE_PROVEN. Historical document numbers above are source evidence, not approved current previews.

## 11. BLOCKERS

Bent-channel source contradiction; missing independently proven historical Q for square examples; missing geometry entries/aliases; no approved implemented channel/square Stage4 rule. Working mark catalogue now exists and is no longer an absence blocker. Source proof is not implementation. Existing filtered XLS input integration remains a later Stage5 concern.

## 12. USER DECISIONS REQUIRED

Need authoritative actual controlled count and explanation for act8 Пп1-13 versus matching VOM8 (7800×6 versus11items). Do not choose Q13 or edit VOM by inference. For square, provide historical controlled-count evidence/registry snapshots for additional independent examples or a separate technically verified inspection-length method. User approval alone must not be described as historical source proof. No repeated question about current Q or current catalogue.

## 13. CHANGED FILES

Only audit scripts/JSON/this report/manifest created under validation/stage5_uzk_source_proof_20261007. No business file changed;{len(before)} protected files unchanged. No new УЗК.docx or PDF. DEV checkpoint unchanged. One quantity-extraction serialization attempt failed due to the PowerShell automatic Matches variable; audit helper variable corrected and read-only extraction rerun; workbook bytes unchanged. No product code fix occurred.

## 14. NOT_RUN

After conflict/quantity STOP, no expansion to other projects/339docs, no external hyperlink target access, no new native Word rendering. Current square/bent calculation previews, Stage4 edits/manual profiles/normalization changes, import, registry mutation, source repair, VBA, web, productionВЗМК,АКЗ,quality document,PDF generation,EXE,deploy:NOT_RUN.

## 15. READY TO IMPLEMENT STAGE 5 — NO

Exact16П historical writing rule proven within bounded size. Bent typeCONTRADICTORY_EVIDENCE; squareINSUFFICIENT_EVIDENCE. Read-only source audit stops here. No implementation begun.
'''
# Derive detail numbers rather than repeating an unverified number in prose.
kmd=read(OUT/'project_kmd_mark_evidence.json');details={r['mark']:r['detail'] for r in kmd}
report=report.replace('details286/293',f"details{details['Бп2-6']}/{details['Бп2-13']}")
(OUT/'AUDIT_REPORT.md').write_text(report,encoding='utf-8')
manifest={'created_at_utc':datetime.now(timezone.utc).isoformat(),'status':summary['status'],'protected_files_checked':len(before),'protected_changes':[],'working_catalog_sha256':catalog['sha256'],'evidence_sha256':{p.name:sha(p) for p in sorted(OUT.iterdir()) if p.is_file() and p.name!='EVIDENCE_MANIFEST.json'}}
save('EVIDENCE_MANIFEST.json',manifest)
assert all(sha(OUT/n)==s for n,s in manifest['evidence_sha256'].items())
print(json.dumps({'protected':len(before),'changed':0,'documents':len(reports),'total_mark_rows':summary['project_document_mark_rows'],'target_examples':len(targets),'rules':rules,'legacy_problem_rows':[r for r in legacy if r['cardinality_mismatch'] or any(s['status']!='MATCH' for s in r['segments'])]},ensure_ascii=False))
