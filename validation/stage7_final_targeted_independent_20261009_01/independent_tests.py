"""New independent tests: never import or execute the producer's test suite."""
import sys
sys.dont_write_bytecode=True
from copy import deepcopy
from unittest.mock import patch,Mock
from pathlib import Path
import json,traceback
from PIL import Image,ImageChops
from review_checks import *
sys.path.insert(0,str(ROOT))
from stage7_quality import core,word,provenance as P,presentation_revision as R,sources as S

M=read(MANIFEST);Q=read(RECEIPT);RESULTS=[];NEG=[]
FIXTURES=OUT/'isolated_fixtures_final';FIXTURES.mkdir(exist_ok=False)
ACTS='Акт ВИК № 9/2026 от 29.09.2026 г., Акт ВИК № 9/2026-1 от 29.09.2026 г., Акт УЗК № 9/2026 от 01.10.2026 г., Сертификат качества антикоррозийной обработки № 9/2026 от 05.10.2026 г.'
STORAGE='Документы о качестве на металлопрокат, сварочные материалы и лакокрасочные материалы хранятся на предприятии'
qc,mc=map(cells,(xml(NEW),xml(MANUAL)))
def check(name,fn):
    try:
        value=fn();RESULTS.append({'name':name,'status':'PASS','evidence':value})
        print('PASS '+name,flush=True)
    except Exception as exc:
        RESULTS.append({'name':name,'status':'FAIL','error':repr(exc),'traceback':traceback.format_exc()})
        print('FAIL '+name+' '+repr(exc),flush=True)
def require(condition,message='assertion failed'):
    if not condition:raise AssertionError(message)
def reseal(m):
    m.pop('seal',None);m['seal']=P.digest(m);return m
def negative(name,fn,expected):
    try:fn()
    except P.Stage7Error as exc:
        require(expected in str(exc),f'expected {expected}; actual {exc}')
        NEG.append({'fixture':name,'expected':expected,'actual':str(exc),'status':'REJECTED'});return str(exc)
    raise AssertionError('Invalid fixture accepted: '+name)
def manifest_fixture(name,mutate,expected):
    m=deepcopy(M);mutate(m);reseal(m)
    p=FIXTURES/(name+'.json');p.write_text(json.dumps(m,ensure_ascii=False,indent=2),encoding='utf8')
    spy=Mock(side_effect=AssertionError('writer must not run'))
    with patch.object(word,'write_quality',spy):
        result=negative(name,lambda:core.generate_quality(p,FIXTURES/(name+'_output')),expected)
    require(spy.call_count==0);require(not (FIXTURES/(name+'_output')).exists());return result
def charformat(cell):
    normalized=[]
    for p in formatting(cell):
        chars=[]
        for run in p['runs']:chars.extend((c,run['rPr']) for c in run['text'])
        normalized.append((p['pPr'],chars))
    return normalized
def presentation13():
    require([text(p) for p in qc[0][50][0].findall('w:p',NS)]==[ACTS])
    require(word.acts_paragraph(M,'basis_point13')==ACTS)
    require(charformat(qc[0][50][0])==charformat(mc[0][50][0]))
    return {'paragraphs':1,'text_exact_manual':True,'all_character_and_paragraph_properties_equal':True}
def presentation14():
    require([text(p) for p in qc[0][53][0].findall('w:p',NS)]==[ACTS,STORAGE])
    require(word.acts_paragraph(M,'attachment_point14')==ACTS)
    require(charformat(qc[0][53][0])==charformat(mc[0][53][0]))
    require(read(OUT/'OOXML_REVIEW.json')['point14_explanation_equal'])
    return {'paragraphs':2,'acts_equal_point13':True,'storage_exact':True,'all_character_and_paragraph_properties_equal':True}
def no_admin():
    content=text(qc[0][50][0])+text(qc[0][53][0])
    for phrase in ('ТЕСТ','SAMPLE_ONLY','АКЗ №0','Подпись','3 назначенные','Приложение','АЦСТ','экз.'):
        require(phrase not in content,phrase)
    return {'admin_insertions':0,'separate_akz_appendix':False,'invented_copy_counts':False}
def appendix_still_bound():
    require(len(M['records'])==6)
    act,app=M['records'][-2:]
    require((act['type'],app['type'])==('AKZ','AKZ_APPENDIX'))
    require(app['parent_id']==act['document_id'] and app['pair_id']==act['pair_id']==read(S.PAIR)['act_id'])
    require(sha(act['path'])==act['sha256'] and sha(app['path'])==app['sha256'])
    require(Q['inputs'][app['path']]==app['sha256']);require(app['type'] not in word.references(M,'attachment_point14'))
    return {'mandatory_pair_verified':True,'akz_sha256':act['sha256'],'appendix_sha256':app['sha256'],'pair_id':act['pair_id']}
def protected12():
    x=read(OUT/'OOXML_REVIEW.json');require(x['point12_canonical_xml_equal'])
    for key in ('.//w:numPr_equal_first','.//w:sectPr_equal_first','.//w:tblPr_equal_first','.//w:trPr_equal_first'):require(x[key])
    a=Image.open(OUT/'native_range_recheck/first/point12.png').convert('RGB');b=Image.open(OUT/'native_range_recheck/new/point12.png').convert('RGB')
    require(a.size==b.size)
    diff=ImageChops.difference(a,b);count=sum(pixel!=(0,0,0) for pixel in diff.get_flattened_data())
    diff.save(OUT/'POINT12_INDEPENDENT_PIXEL_DIFF.png')
    require(count==0,f'point12 direct native range: {count} differing pixels')
    n=read(OUT/'NATIVE_VERIFIED_REVIEW.json');f=n['first']['ranges']['point12'];q=n['new']['ranges']['point12']
    require(f['text']==q['text'])
    require(f['top_pt']==q['top_pt'] and f['end_pt']==q['end_pt'] and f['page']==q['page']==1)
    return {'xml_rows':x['point12_rows_1based'],'canonical_xml_equal':True,'native_text_equal':True,'native_coordinates_equal':True,'native_range_png_dimensions':a.size,'differing_pixels':count}
def point12_format_mutation():
    f=xml(FIRST);q=xml(NEW);r=q.findall('w:body/w:tbl',NS)[0].findall('w:tr',NS)[40]
    p=r.find('.//w:pPr',NS)
    E.SubElement(p,'{'+NS['w']+'}keepNext')
    return negative('point12_format_in_memory',lambda:word.structural_check(f,q),'WORD_CONSTANT_CELL_CHANGED')
def contract():
    require(text(qc[0][11][2])=='Договор №')
    require(text(qc[0][11][3]).strip()=='06/26 от 03.06.2026')
    require(read(OUT/'OOXML_REVIEW.json')['changed_cells_vs_previous']==[[0,11,3],[0,50,0],[0,53,0]])
    require(word.contract_value('  №  88/31 от 02.01.2031  ')=='88/31 от 02.01.2031')
    negative('double_number_prefix',lambda:word.contract_value('№ № 88'),'CONTRACT_VALUE_INVALID')
    return {'visible':'Договор № 06/26 от 03.06.2026','only_expected_3_cells_changed':True}
def dynamic():
    m=deepcopy(M);selected=[r for r in m['records'] if r['type'] in ('VIK','UZK','AKZ')]
    values=[('51/2032','02.01.2032'),('51/2032-2','03.01.2032'),('UZ-77','04.01.2032'),('AK-88','05.01.2032')]
    for r,(number,date) in zip(selected,values):r.update(number=number,date=date,display_title='POISONED DISPLAY TITLE')
    expected='Акт ВИК № 51/2032 от 02.01.2032 г., Акт ВИК № 51/2032-2 от 03.01.2032 г., Акт УЗК № UZ-77 от 04.01.2032 г., Сертификат качества антикоррозийной обработки № AK-88 от 05.01.2032 г.'
    for usage in ('basis_point13','attachment_point14'):
        require(word.acts_paragraph(m,usage)==expected)
        cell=deepcopy(qc[0][50][0]);word.list_cell(cell,[word.acts_paragraph(m,usage)])
        require(text(cell)==expected)
    return {'independent_distinct_numbers_and_dates':values,'display_title_not_used':True,'scope':'presenter in memory, not a new approved assembly'}
def false_reference():
    m=deepcopy(M)
    m['records'].append({'type':'HISTORIC_CERTIFICATE','usage':['basis_point13','attachment_point14'],'display_title':'FALSE','number':'FALSE','date':'01.01.1900'})
    require(word.acts_paragraph(m,'basis_point13')==ACTS and word.acts_paragraph(m,'attachment_point14')==ACTS)
    require([r['type'] for r in M['records']]==['QUALITY_APPENDIX_1','VIK','VIK','UZK','AKZ','AKZ_APPENDIX'])
    return {'no_historical_reference':True,'all_4_visible_acts_match_verified_records':True}
def test_status():
    require(word.BANNER in text(xml(NEW)))
    require(M['mode']=='DEV_TEST_ONLY' and M['official_issue'] is False and Q['official_issue'] is False and Q['results_confirmed'] is False and Q['status']=='TEST')
    require(M['records'][3]['status']=='SAMPLE_ONLY' and all(r['official_confirmed'] is False for r in M['records']))
    scope=M['akz_scope'];require(len(scope['assigned_marks'])==3 and scope['quantity']==3 and scope['shipment_quantity']==288 and scope['actual_measurements_confirmed'] is False and scope['variant']==0)
    require(read(OUT/'OOXML_REVIEW.json')['signature_cells_equal_old'])
    require(read(OUT/'OOXML_REVIEW.json')['drawings']==0)
    return {'uzk_status':'SAMPLE_ONLY','results_confirmed':False,'official_issue':False,'akz_scope':scope,'added_signatures_or_seals':0}
def native():
    n=read(OUT/'NATIVE_VERIFIED_REVIEW.json');q=n['new'];m=n['manual']
    require(q['pages']==2 and q['sections']==1 and q['tables']==2 and q['ReadOnly'] and not q['OpenAndRepair'] and not q['Save'])
    require(abs(q['width_pt']-595.3)<0.1 and abs(q['height_pt']-841.9)<0.1)
    require([p['label'] for p in q['lists'] if p['label'].rstrip('.').isdigit()]==[str(i)+'.' for i in range(1,15)])
    for k in ('point13','point14','storage'):
        for property in ('text','font','paragraph','page','top_pt','style'):require(q['ranges'][k][property]==m['ranges'][k][property],f'{k} {property}')
    require(q['shapes']==q['inline_shapes']==q['revisions']==0)
    require(q['sha_before']==q['sha_after']==EXPECTED_NEW)
    return {'pages':2,'a4_pt':[q['width_pt'],q['height_pt']],'logical_points':14,'manual_effective_font_paragraph_and_positions_equal':True,'read_only_no_repair_no_save':True}
def manifest_receipt():
    core.verify_manifest(M)
    reference={'path':str(RECEIPT),'sha256':sha(RECEIPT)}
    verified=P.verify_receipt(reference,M['assembly_run_id'],S.PINS[S.REGISTRY],str(NEW),EXPECTED_NEW,'TEST')
    require(verified==Q and Q['code']==core.code_bindings()==read(FIX/'PRODUCT_AFTER_FINAL_TESTS.json'))
    require(Q['inputs'][str(MANIFEST)]==sha(MANIFEST))
    events={k:read(Q[k]['path']) for k in ('start','result','validation')}
    for k,e in events.items():
        require(sha(Q[k]['path'])==Q[k]['sha256']);require(e['operation_id']==Q['operation_id'] and e['assembly_run_id']==M['assembly_run_id'])
    require(events['start']['expected_output']==str(NEW) and events['result']['sha256']==EXPECTED_NEW and events['validation']['checks']['status']=='PASS')
    require(events['start']['code']==Q['code'] and events['start']['inputs']==Q['inputs'])
    for name,(old,current) in R.VERSIONS.items():require(Q['inputs'][str(R.ARCHIVE/name)]==old and sha(ROOT/'stage7_quality'/name)==current)
    save('VERIFIED_CHAIN.json',{'receipt_sha256':sha(RECEIPT),'manifest_sha256':sha(MANIFEST),'assembly_run_id':M['assembly_run_id'],'operation_id':Q['operation_id'],'status':Q['status'],'official_issue':Q['official_issue'],'results_confirmed':Q['results_confirmed'],'code_files':len(Q['code']),'inputs':len(Q['inputs']),'events':{k:Q[k] for k in events},'records':M['records']})
    return {'current_code_bindings':len(Q['code']),'receipt_sha256':sha(RECEIPT),'manifest_sha256':sha(MANIFEST),'event_chain':'START -> RESULT -> VALIDATION','historical_sha_archives_strictly_bound':True}
def archived_code():
    evidence=R.verify_producer_code(M['code']);require(len(evidence)==3)
    unknown=dict(M['code']);unknown[str(ROOT/'stage7_quality/word.py')]='0'*64
    negative('unknown_old_stage7_sha',lambda:R.verify_producer_code(unknown),'INPUT_OR_CODE_CHANGED')
    unknown=dict(M['code']);unknown[str(ROOT/'passport_app/vom.py')]='0'*64
    negative('unchanged_stage1_sha',lambda:R.verify_producer_code(unknown),'INPUT_OR_CODE_CHANGED')
    archive=FIXTURES/'code_archive';archive.mkdir()
    for name in R.VERSIONS:(archive/name).write_bytes((R.ARCHIVE/name).read_bytes())
    with (archive/'word.py').open('ab') as s:s.write(b'\n# independent invalid fixture\n')
    with patch.object(R,'ARCHIVE',archive):negative('altered_archive',lambda:R.verify_producer_code(M['code']),'INPUT_OR_CODE_CHANGED')
    original=R.sha
    with patch.object(R,'sha',side_effect=lambda p:'0'*64 if Path(p)==ROOT/'stage7_quality/word.py' else original(p)):
        negative('unapproved_current_revision',lambda:R.verify_producer_code(M['code']),'UNAPPROVED_STAGE7_CODE_REVISION')
    return {'allowed_exact_transitions':3,'unknown_old_rejected':True,'stage1_strict':True,'altered_archive_rejected':True,'unapproved_current_rejected':True}
def receipt_official():
    r=deepcopy(Q);r['official_issue']=True
    f=FIXTURES/'official_receipt.json';f.write_text(json.dumps(r,ensure_ascii=False,indent=2),encoding='utf8')
    return negative('official_receipt',lambda:P.verify_receipt({'path':str(f),'sha256':sha(f)},M['assembly_run_id'],S.PINS[S.REGISTRY],str(NEW),EXPECTED_NEW,'TEST'),'RECEIPT_IDENTITY_MISMATCH')
def presenter_official():
    m=deepcopy(M);m['official_issue']=True;output=FIXTURES/'OFFICIAL_MUST_NOT_EXIST.docx'
    negative('official_writer',lambda:word.write_quality(m,output),'OFFICIAL_FORBIDDEN');require(not output.exists())
    m=deepcopy(M);m['records'][3]['status']='OFFICIAL'
    return negative('official_visible_record',lambda:word.acts_paragraph(m,'basis_point13'),'REFERENCE_STATUS_INVALID')
def preserved_sources():
    h=read(OUT/'HASH_REVIEW.json');require(h['protected_count']==1264 and h['product_count']==103)
    for k in ('protected_current_mismatches','product_current_mismatches','evidence_current_mismatches','stage1_6_changes','extra_product_source_files'):require(not h[k],k)
    require(h['new_distinct_from_old'] and h['old_sha_matches_previous_review'] and h['old_manifest_matches_previous_review'] and h['supplied_diff_equal'])
    return {'protected_files':1264,'product_files':103,'old_output_manifest_and_receipts_preserved':True,'expected_product_changes':4}

check('01 Point 13 exact manual text, paragraph, all characters formatting',presentation13)
check('02 Point 14 same acts, exact storage and explanation, all characters formatting',presentation14)
check('03 No administrative inserts or separate AKZ-app in Word',no_admin)
check('04 Mandatory AKZ pair still bound by SHA and parent/pair id',appendix_still_bound)
check('05 Point 12 complete XML and independent native pixel comparison',protected12)
check('06 Point 12 isolated formatting mutation rejected',point12_format_mutation)
check('07 Contract and other fields unchanged outside the three cells',contract)
check('08 Dynamic distinct numbers/dates, poisoned display titles ignored',dynamic)
check('09 No false links or historical references',false_reference)
check('10 TEST/SAMPLE_ONLY, three-mark AKZ, no signatures or seals',test_status)
check('11 Native Word physical pages, points, effective formatting and coordinates',native)
check('12 New output manifest/receipt/current code and event chain',manifest_receipt)
check('13 Missing VIK record rejected before writer',lambda:manifest_fixture('missing_vik',lambda m:m['records'].pop(1),'SELECTED_RECORDS_MISMATCH_OR_MISSING'))
check('14 Missing UZK file rejected before writer',lambda:manifest_fixture('missing_uzk_file',lambda m:m['records'][3].update(path=str(FIXTURES/'absent_uzk.docx')),'INPUT_OR_CODE_CHANGED'))
check('15 Incomplete AKZ pair (no appendix record) rejected before writer',lambda:manifest_fixture('missing_akz_appendix',lambda m:m['records'].pop(5),'SELECTED_RECORDS_MISMATCH_OR_MISSING'))
check('16 Incomplete AKZ pair (missing appendix file) rejected before writer',lambda:manifest_fixture('missing_akz_file',lambda m:m['records'][5].update(path=str(FIXTURES/'absent_akz.xlsx')),'INPUT_OR_CODE_CHANGED'))
check('17 Incorrect pair id rejected before writer',lambda:manifest_fixture('wrong_pair',lambda m:m['records'][5].update(pair_id='unrelated'),'AKZ_PAIR_OR_VARIANT_MISMATCH'))
check('18 Official manifest rejected before writer',lambda:manifest_fixture('official_manifest',lambda m:m.update(official_issue=True),'MANIFEST_SCOPE_OR_OFFICIAL_FORBIDDEN'))
check('19 Promoted UZK status rejected before writer',lambda:manifest_fixture('official_uzk_status',lambda m:m['records'][3].update(status='TEST'),'SELECTED_RECORDS_MISMATCH_OR_MISSING'))
check('20 Wrong referenced date rejected before writer',lambda:manifest_fixture('wrong_uzk_date',lambda m:m['records'][3].update(date='01.10.2031'),'SELECTED_RECORDS_MISMATCH_OR_MISSING'))
check('21 Only three archived code transitions, no broad hash bypass',archived_code)
check('22 Official receipt rejected',receipt_official)
check('23 Official writer and presenter rejected',presenter_official)
check('24 Protected sources, current product and old artifacts unchanged',preserved_sources)
save('NEGATIVE_FIXTURES.json',NEG)
summary={'status':'PASS' if all(t['status']=='PASS' for t in RESULTS) else 'FAIL','total':len(RESULTS),'passed':sum(t['status']=='PASS' for t in RESULTS),'failed':sum(t['status']=='FAIL' for t in RESULTS),'tests':RESULTS,'negative_cases':len(NEG),'document_generations':{'VIK':0,'UZK':0,'AKZ':0,'QUALITY':0},'official_quality_guard_test_invocations':1,'scope':'independent targeted suite; producer suite not executed'}
save('INDEPENDENT_TEST_RESULTS.json',summary)
print(json.dumps({k:v for k,v in summary.items() if k!='tests'},ensure_ascii=False,indent=2))
sys.exit(0 if summary['status']=='PASS' else 1)
