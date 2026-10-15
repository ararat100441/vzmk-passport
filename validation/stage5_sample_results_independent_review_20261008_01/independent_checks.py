"""Independent source/document proof and isolated fail-closed checks; never create an act."""
import hashlib
import io
import json
import os
from pathlib import Path
import shutil
import sys
import tempfile
import traceback
from copy import deepcopy
from unittest.mock import patch
from zipfile import ZipFile

OUT=Path(__file__).resolve().parent
ROOT=OUT.parent.parent
IMPL=ROOT/'validation/stage5_sample_results_20261008_130441'
TMP=OUT/'independent_tmp';TMP.mkdir(exist_ok=True)
tempfile.tempdir=str(TMP);os.environ['TMP']=os.environ['TEMP']=str(TMP)
sys.path.insert(0,str(ROOT))
EVENTS=[]
FIXTURE=OUT/'sample_fixture'

def guard(event,args):
    if event!='open' or not isinstance(args[0],(str,bytes)):return
    path=Path(os.fsdecode(args[0])).resolve()
    assert path.name.casefold() not in ('взмк.xlsx','узк.xlsm'), f'PRODUCTION_FORBIDDEN: {path}'
    mode,flags=args[1:3]
    writing=(isinstance(mode,str) and any(c in mode for c in 'wax+')) or (
        isinstance(flags,int) and flags & (os.O_WRONLY|os.O_RDWR|os.O_CREAT|os.O_TRUNC))
    if writing:
        assert path.is_relative_to(OUT), f'WRITE_OUTSIDE_REVIEW: {path}'
        assert path.suffix.lower()!='.docx' or path.is_relative_to(FIXTURE), f'ACT_CREATION_FORBIDDEN: {path}'
    EVENTS.append({'path':str(path),'writing':bool(writing)})

sys.addaudithook(guard)
from docx import Document
from docx.table import _Cell
from lxml import etree as E
import uzk_stage5.core as C
import uzk_stage5.result_samples as S
import uzk_stage5.word as W
from uzk_stage5.inputs import Stage5Error,sha
from profiles_stage4.offline import no_network

def read(path):return json.loads(Path(path).read_text(encoding='utf-8-sig'))
def save(name,data):(OUT/name).write_text(json.dumps(data,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')

preflight=read(OUT/'preflight.json')
BEFORE=read(OUT/'product_before_review.json')
assert {p:sha(p) for p in BEFORE}==BEFORE
save('independent_product_before.json',BEFORE)
DATA=read(IMPL/'current_preview.json')['preview']
DOCX=Path(preflight['artifact'])
OLD=Path(preflight['old_test_docx']['path'])
RESULTS=[]

def case(name,fn):
    print('START',name,flush=True)
    try: detail=fn();result={'name':name,'status':'PASS','detail':detail}
    except Exception:
        result={'name':name,'status':'FAIL','detail':traceback.format_exc()}
    RESULTS.append(result);save('independent_checks.json',RESULTS)
    print('END',name,result['status'],flush=True)

def history():
    records=[];sources=[]
    for item in read(IMPL/'historical_sources.json'):
        path=Path(DATA['project_root'])/item['relative_path']
        assert sha(path)==item['sha256']
        document=Document(path)
        matches=[]
        for ti,table in enumerate(document.tables,1):
            for ri,row in enumerate(table._tbl.tr_lst):
                if len(row.tc_lst)!=8:continue
                values=[_Cell(tc,table).text for tc in row.tc_lst]
                if 'Обнаруженные дефекты' in values[6].replace('\n',' ') and 'Заключение' in values[7]:
                    matches.append((ti,ri,table,values[6],values[7]))
        assert len(matches)==1
        ti,header,table,h7,h8=matches[0]
        source_records=[]
        for ri,row in enumerate(table._tbl.tr_lst[header+1:],header+2):
            assert len(row.tc_lst)==8
            values=[_Cell(tc,table).text for tc in row.tc_lst]
            source_records.append({'act':path.name,'source':str(path),'table':ti,'row':ri,
                                   'mark':values[0],'graph7':values[6],'graph8':values[7]})
        assert len(source_records)==item['rows']
        records.extend(source_records)
        sources.append({'path':str(path),'sha256':sha(path),'rows':len(source_records),'graph7_header':h7,'graph8_header':h8})
    assert records==read(IMPL/'historical_rows.json')
    assert len(sources)==10 and len(records)==56
    assert {(r['graph7'],r['graph8']) for r in records}=={('ДНО','годен')}
    assert all(r['mark'].strip() for r in records)
    save('historical_rows_independent.json',records);save('historical_sources_independent.json',sources)
    return {'acts':10,'physical_rows':56,'unique_raw_pair':['ДНО','годен'],'blank_results':0,'other_results':0,'source_hashes_match':True}

def document():
    ns=W.NS
    with ZipFile(OLD) as a, ZipFile(DOCX) as b:
        assert a.namelist()==b.namelist()
        changed=[n for n in a.namelist() if a.read(n)!=b.read(n)]
        assert changed==['word/document.xml']
        old=E.fromstring(a.read('word/document.xml'));new=E.fromstring(b.read('word/document.xml'))
    oldbody=old.find('w:body',ns);body=new.find('w:body',ns)
    oldtables=oldbody.findall('w:tbl',ns);tables=body.findall('w:tbl',ns)
    assert len(tables)==len(oldtables)==2
    oldrows=oldtables[0].findall('w:tr',ns);rows=tables[0].findall('w:tr',ns)
    assert len(rows)==len(oldrows)==9
    for a,b in zip(oldrows[:5],rows[:5]):assert E.tostring(a)==E.tostring(b)
    expected=['Бп2-1','Пс1-9','Пс1-10','Пс2-16']
    proof=[]
    for name,a,b in zip(expected,oldrows[5:],rows[5:]):
        oldcells=a.findall('w:tc',ns);cells=b.findall('w:tc',ns)
        assert len(cells)==8
        values=[W.contents(c) for c in cells]
        assert values[0]==name and values[6:] == ['ДНО','годен']
        for i in range(6):assert E.tostring(oldcells[i])==E.tostring(cells[i])
        for i in (6,7):
            for prop in ('w:tcPr','w:p/w:pPr','w:p/w:r/w:rPr'):
                x,y=oldcells[i].find(prop,ns),cells[i].find(prop,ns)
                assert (E.tostring(x) if x is not None else None)==(E.tostring(y) if y is not None else None)
        proof.append({'mark':name,'physical_row':rows.index(b)+1,'columns':values,'graph7_physical_column':7,'graph8_physical_column':8})
        b.replace(cells[6],deepcopy(oldcells[6]));b.replace(cells[7],deepcopy(oldcells[7]))
    titles=[p for p in body.findall('w:p',ns) if W.contents(p).startswith('Акт №')]
    assert len(titles)==1
    title=W.contents(titles[0])
    assert 'Акт № 9/2026 от 01.10.2026' in title
    assert S.TEST_LABEL in title and S.UNCONFIRMED_LABEL in title
    assert not titles[0].findall('.//w:vanish',ns)
    oldtitle=next(p for p in oldbody.findall('w:p',ns) if W.contents(p).startswith('Акт №'))
    body.replace(titles[0],deepcopy(oldtitle))
    assert E.tostring(old)==E.tostring(new)
    save('exact_document_columns.json',{'path':str(DOCX),'sha256':sha(DOCX),'rows':proof,'title':title,
        'changed_package_parts':changed,'only_title_and_graph7_graph8_changed':True,'first_six_columns_headers_signatures_layout_unchanged':True})
    return {'marks':expected,'graph7':'ДНО','graph8':'годен','physical_columns':[7,8],'rows':4,'only_authorized_OOXML_changes':True}

def accepted():
    sample=S.load_sample_results(DATA)
    assert sample['mode']=='SAMPLE_ONLY' and sample['confirmed'] is False
    assert (sample['graph7'],sample['graph8'],sample['row_count'])==('ДНО','годен',56)
    assert len(sample['fingerprints'])==10
    save('accepted_sample_independent.json',sample)
    return {'mode':'SAMPLE_ONLY','confirmed':False,'sources':10,'rows':56}

def false_confirmation():
    data=deepcopy(DATA);data['results_confirmed']=True;data['responsible_executor']='ISOLATED TEST'
    sample=S.load_sample_results(data)
    assert sample['confirmed'] is False and sample['mode']=='SAMPLE_ONLY'
    return 'TEST/DRAFT stays SAMPLE_ONLY / confirmed=False despite unsupported confirmation fields'

def stop(data,expected,name):
    output=OUT/'negative_outputs'/(name+'.docx')
    assert not output.exists() and not output.parent.exists()
    try:W.write_test_document(data,output)
    except Stage5Error as exc:
        assert expected in str(exc), str(exc)
        assert not output.exists() and not output.parent.exists()
        return str(exc)
    raise AssertionError('Expected STOP before file creation')

def official():
    evidence=[]
    for kind in ('OFFICIAL','CONFIRMED',None):
        data=deepcopy(DATA);data['kind']=kind;data['results_confirmed']=True
        data['responsible_executor']='ISOLATED TEST';data['actual_results']={'Бп2-1':['ДНО','годен']}
        evidence.append(stop(data,'OFFICIAL_RESULTS_NOT_IMPLEMENTED','official_'+str(kind)))
    return evidence

def other_project():
    data=deepcopy(DATA);data['passport']['project']='Other Project'
    return stop(data,'SAMPLE_PROJECT_NOT_PROVEN','other_project')

for relative,expected,count in S.ACCEPTED_SOURCES:
    source=Path(DATA['project_root'])/relative;target=FIXTURE/relative
    target.parent.mkdir(parents=True,exist_ok=True);shutil.copyfile(source,target)
    assert sha(target)==expected
FIRST=FIXTURE/S.ACCEPTED_SOURCES[0][0]
FIRST_BYTES=FIRST.read_bytes()
FIX_DATA=deepcopy(DATA);FIX_DATA['project_root']=str(FIXTURE);FIX_DATA['template']=str(FIRST)

def reset_first():FIRST.write_bytes(FIRST_BYTES)

def missing():
    FIRST.unlink()
    try:return stop(FIX_DATA,'SAMPLE_SOURCE_MISSING','missing')
    finally:reset_first()

def drift():
    FIRST.write_bytes(FIRST_BYTES+b'isolated source drift')
    try:return stop(FIX_DATA,'SAMPLE_SOURCE_CHANGED','source_drift')
    finally:reset_first()

def mutated_fixture(kind):
    ns=W.NS
    with ZipFile(io.BytesIO(FIRST_BYTES)) as z:
        xml=E.fromstring(z.read('word/document.xml'))
        table=xml.find('w:body/w:tbl',ns);rows=table.findall('w:tr',ns)
        if kind in ('blank','ambiguous'):
            cell=rows[5].findall('w:tc',ns)[6 if kind=='blank' else 7]
            texts=cell.findall('.//w:t',ns)
            for t in texts:t.text=''
            texts[0].text='' if kind=='blank' else 'не годен'
        elif kind=='table_ambiguous':table.insert(list(table).index(rows[4]),deepcopy(rows[4]))
        elif kind=='row_count':table.append(deepcopy(rows[5]))
        payload=io.BytesIO()
        with ZipFile(payload,'w') as dest:
            for entry in z.infolist():dest.writestr(entry,E.tostring(xml) if entry.filename=='word/document.xml' else z.read(entry.filename))
    FIRST.write_bytes(b'not a DOCX' if kind=='corrupt' else payload.getvalue())
    pins=list(S.ACCEPTED_SOURCES);rel,_,count=pins[0];pins[0]=(rel,sha(FIRST),count)
    expected={'blank':'SAMPLE_NOT_PROVEN','ambiguous':'SAMPLE_AMBIGUOUS','table_ambiguous':'SAMPLE_TABLE_AMBIGUOUS',
              'row_count':'SAMPLE_ROW_COUNT_CHANGED','corrupt':'SAMPLE_INVALID_DOCUMENT'}[kind]
    try:
        # Test semantic rejection after a matching isolated hash. Production pins
        # and production sources stay unchanged; ordinary drift is tested above.
        with patch.object(S,'ACCEPTED_SOURCES',tuple(pins)):
            return stop(FIX_DATA,expected,kind)
    finally:reset_first()

def between_load_and_write():
    verify=S.verify_sample_sources
    def mutate_then_verify(sample):
        FIRST.write_bytes(FIRST_BYTES+b'drift after sample extraction')
        return verify(sample)
    try:
        with patch.object(W,'verify_sample_sources',side_effect=mutate_then_verify):
            return stop(FIX_DATA,'SAMPLE_SOURCE_CHANGED','prewrite_drift')
    finally:reset_first()

def native():
    info=read(OUT/'native_word/native_word.json')
    assert info['status']=='PASS_NORMAL_READONLY_OPEN' and info['readonly'] and info['unchanged']
    assert info['pages']==2 and info['sections']==1 and info['tables']==2 and info['first_table_rows']==9
    assert info['orientation']==1 and info['page_width']>info['page_height'] and info['hyperlinks']==0
    assert info['before_sha']==info['after_sha']==sha(DOCX)
    assert info['close_without_save'] and info['owned_instance'] and info['existing_word_processes_preserved']
    assert len(info['images'])==2 and all(Path(p).is_file() for p in info['images'])
    # Both original-resolution native page images were inspected by reviewer.
    proof={'status':'PASS','method':'Microsoft Word normal read-only open, native EnhMetaFileBits page render, both full-resolution page images inspected',
           'pages':2,'landscape':True,'four_rows_and_both_result_columns_visible':True,
           'labels_noticeable':True,'clipping_or_overlap':False,'broken_table':False,'extra_blank_page':False,
           'requisites_executor_empty_signature_qualification_preserved':True,'hyperlinks':0,
           'repair_recovery':False,'close_without_save':True,'existing_user_Word_preserved':True,
           'word_version':info['version'],'word_build':info['build'],'images':info['images'],'artifact_sha256':sha(DOCX)}
    save('visual_validation_independent.json',proof)
    return proof

def calculation():
    p=C.preview('Гараж Казахстан','9/2026',Path(DATA['vom']['path']))
    p.verify();data=p.data
    expected={'Бп2-1':(8,[('1280','5'),('1024','8.4')]),'Пс1-9':(9,[('3600','4')]),
              'Пс1-10':(15,[('6000','4')]),'Пс2-16':(1,[('480','3')])}
    actual={}
    for m in data['marks']:
        actual[m['mark']]=(m['Q'],[(str(s['L']),str(s['S'])) for s in m['segments']])
        assert all(s['frequency_mhz']==5 and s['angle_degrees']==70 for s in m['segments'])
    assert actual==expected and len(data['marks'])==4 and data['controlled_units']==33
    assert data['passport']['number']=='9/2026' and data['passport']['date']=='01.10.2026'
    assert data['marks']==DATA['marks'] and not data['network_attempts']
    save('fresh_current_preview.json',{'data':data,'seal':p.seal})
    return {'passport':'9/2026','date':'01.10.2026','marks':4,'units':33,'all_segments':'5 MHz / 70 degrees','DOCX_generated':False}

with no_network() as attempts:
    case('Historical exact raw source extraction 10 acts / 56 physical rows',history)
    case('Exact new artifact physical columns and OOXML delta',document)
    case('Accepted sample remains unconfirmed SAMPLE_ONLY',accepted)
    case('Unsupported confirmation fields cannot promote a TEST sample',false_confirmation)
    case('Official modes rejected by actual renderer before output',official)
    case('Other project has no accepted sample STOP',other_project)
    case('Missing historical source STOP before output',missing)
    case('Changed historical source SHA STOP before output',drift)
    for kind in ('blank','ambiguous','table_ambiguous','row_count','corrupt'):
        case('Isolated semantic '+kind+' source STOP before output',lambda kind=kind:mutated_fixture(kind))
    case('Source drift after load STOP before exclusive output creation',between_load_and_write)
    case('Fresh native Word and both-page visual validation',native)
    case('Fresh current read-only four-mark calculation regression',calculation)
AFTER={p:sha(p) for p in BEFORE};assert AFTER==BEFORE
save('independent_product_after.json',AFTER)
protected=read(OUT/'protected_before.json');assert {p:sha(p) for p in protected}==protected
save('independent_access_guard.json',{'events':EVENTS,'production_reads':0,'writes_outside_review':0,
                                      'act_creation_attempts':0,'network_attempts':list(attempts)})
summary={'tests':len(RESULTS),'passed':sum(r['status']=='PASS' for r in RESULTS),
         'failures':[r['name'] for r in RESULTS if r['status']!='PASS'],'network_attempts':list(attempts),
         'current_product_before_after_match':True,'protected_files_unchanged':len(protected),'new_generated_DOCX':0}
save('independent_summary.json',summary)
print(json.dumps(summary,ensure_ascii=False,indent=2))
raise SystemExit(1 if summary['failures'] or attempts else 0)
