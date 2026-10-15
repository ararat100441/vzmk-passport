"""Independent targeted failures, only private manifests and fixtures."""
import sys
sys.dont_write_bytecode=True
import json, hashlib, shutil
from pathlib import Path
from copy import deepcopy
from unittest.mock import patch
ROOT=Path(__file__).resolve().parents[2];OUT=Path(__file__).resolve().parent
sys.path.insert(0,str(ROOT))
from stage7_quality.core import verify_manifest,generate_quality
from stage7_quality.provenance import Stage7Error,digest,read,sha,Recorder,frozen
from stage7_quality import sources as S
M=read(ROOT/'validation/stage7_resume_20261009_175000/ASSEMBLY_01/STAGE7_SELECTED_TEST_MANIFEST.json')
FIX=OUT/'negative_fixtures';FIX.mkdir(exist_ok=False)
results=[]
def reseal(m):
    m.pop('seal',None);m['seal']=digest(m);return m
def save(p,d):p.write_text(json.dumps(d,ensure_ascii=False,indent=2),encoding='utf8')
def reject(name,change,expected):
    m=deepcopy(M);change(m);p=FIX/(name+'.json');save(p,reseal(m))
    try:verify_manifest(read(p))
    except Stage7Error as e:
        r={'test':name,'result':'PASS' if expected in str(e) else 'FAIL','actual':str(e),'expected':expected}
    except Exception as e:r={'test':name,'result':'FAIL','actual':repr(e),'expected':expected}
    else:r={'test':name,'result':'FAIL','actual':'ACCEPTED','expected':expected}
    results.append(r);print(json.dumps(r,ensure_ascii=False),flush=True)

verify_manifest(M)
copy=FIX/'changed_registry.xlsx';shutil.copyfile(S.REGISTRY,copy)
initial=sha(copy)
with copy.open('ab') as f:f.write(b'INDEPENDENT_DRIFT')
reject('01_input_sha_changed',lambda m:m['inputs'].update({str(copy):initial}),'INPUT_OR_CODE_CHANGED')
reject('02_required_uzk_missing',lambda m:m['records'].pop(3),'SELECTED_RECORDS_MISMATCH_OR_MISSING')
reject('03_old_vik_registry',lambda m:m['records'][1]['membership'].update(registry_sha256='0'*64),'RECORD_MEMBERSHIP_OR_OFFICIAL_MISMATCH')
reject('04_akz_pair_incomplete',lambda m:m['records'].pop(5),'SELECTED_RECORDS_MISMATCH_OR_MISSING')
reject('05_failed_document',lambda m:m['records'][3].update(status='FAILED_VALIDATION_DO_NOT_USE'),'SELECTED_RECORDS_MISMATCH_OR_MISSING')
reject('06_superseded_document',lambda m:m['records'][4].update(status='SUPERSEDED'),'SELECTED_RECORDS_MISMATCH_OR_MISSING')
reject('07_mixed_source_run_ids',lambda m:m['records'][2].update(source_run_id='other_run'),'WRONG_GENERATOR_OR_RUN')
reject('08_mixed_assembly_run_ids',lambda m:m.update(assembly_run_id='other_assembly'),'RECEIPT_IDENTITY_MISMATCH')
reject('09_official_state',lambda m:m.update(official_issue=True),'MANIFEST_SCOPE_OR_OFFICIAL_FORBIDDEN')
reject('10_claim_confirmed',lambda m:m['records'][3].update(official_confirmed=True),'RECORD_MEMBERSHIP_OR_OFFICIAL_MISMATCH')
reject('11_missing_actual_file',lambda m:m['records'][3].update(path=str(FIX/'missing_uzk.docx')),'INPUT_OR_CODE_CHANGED')
reject('12_akz_broken_parent',lambda m:m['records'][5].update(parent_id='akz1'),'AKZ_PAIR_OR_VARIANT_MISMATCH')
reject('13_akz_variant1',lambda m:m['records'][4].update(variant=1),'AKZ_PAIR_OR_VARIANT_MISMATCH')
reject('14_uzk_not_required',lambda m:m.update(uzk_required=False),'MANIFEST_SCOPE_OR_OFFICIAL_FORBIDDEN')
reject('15_fake_document_reference',lambda m:m['records'][1].update(display_title='Акт ВИК № 999/2026'),'RECORD_PRESENTATION_MISMATCH')

def generation_stop(name,manifest,prepare,expected):
    base=FIX/name;base.mkdir();mp=base/'manifest.json';save(mp,manifest)
    target=base/'output';prepare(target,mp)
    calls=[]
    def sentinel(*args,**kwargs):calls.append(True);raise AssertionError('GENERATOR_REACHED')
    try:
        with patch('stage7_quality.word.write_quality',sentinel):generate_quality(mp,target)
    except Stage7Error as e:status='PASS' if expected in str(e) and not calls else 'FAIL';actual=str(e)
    except Exception as e:status='FAIL';actual=repr(e)
    else:status='FAIL';actual='ACCEPTED'
    r={'test':name,'result':status,'actual':actual,'expected':expected,'generator_calls':len(calls)}
    results.append(r);print(json.dumps(r,ensure_ascii=False),flush=True)

preview=deepcopy(M);late=FIX/'post_preview_input.txt';late.write_text('BEFORE',encoding='utf8')
preview['inputs'][str(late)]=sha(late);preview=reseal(preview);verify_manifest(preview)
late.write_text('AFTER',encoding='utf8')
generation_stop('16_input_after_preview',preview,lambda target,mp:None,'INPUT_OR_CODE_CHANGED')
def existing(target,mp):
    target.mkdir();(target/'9_2026 - Документ о качестве (ТЕСТ).docx').write_bytes(b'EXISTING MUST REMAIN')
generation_stop('17_output_exists',M,existing,'OUTPUT_EXISTS')
assert (FIX/'17_output_exists/output/9_2026 - Документ о качестве (ТЕСТ).docx').read_bytes()==b'EXISTING MUST REMAIN'

# Copy the whole VIK receipt chain, bind it correctly, then replace current registry
# with a historical one. The identity gate must still reject a valid local chain.
def old_chain(m):
    r=m['records'][1];old=read(r['producer_receipt']['path']);base=FIX/'old_vik_chain';base.mkdir()
    old['inputs'][str(S.REGISTRY)]='f'*64
    refs={}
    for key in ['start','result','validation']:
        v=read(old[key]['path'])
        if key=='start':v['inputs'][str(S.REGISTRY)]='f'*64
        if key in ('result','validation'):v['start']=refs['start']
        if key=='validation':v['result']=refs['result']
        p=base/(key+'.json');save(p,v);refs[key]={'path':str(p),'sha256':sha(p)}
    old.update(refs);p=base/'receipt.json';save(p,old);ref={'path':str(p),'sha256':sha(p)}
    for r in m['records'][1:3]:r['producer_receipt']=ref
reject('18_old_vik_actual_receipt',old_chain,'RECEIPT_REGISTRY_MISMATCH')

# A failure after a real isolated callable must never create success receipts.
def recorder_case(name,fn,validator):
    base=FIX/name;base.mkdir();rec=Recorder(base,'independent_test',{},{});rec.register(fn.__module__+'.'+fn.__name__,fn)
    try:rec.run(fn.__module__+'.'+fn.__name__,'TEST',base/'new.txt',{'output':base/'new.txt'},validator)
    except Exception as e:actual=str(e)
    else:actual='ACCEPTED'
    successes=list(base.rglob('STAGE7_OPERATION_RECEIPT.json'));failures=list(base.rglob('FAILURE.json'))
    r={'test':name,'result':'PASS' if not successes and len(failures)==1 and actual!='ACCEPTED' else 'FAIL',
        'actual':actual,'success_receipts':len(successes),'failure_receipts':len(failures)}
    results.append(r);print(json.dumps(r,ensure_ascii=False),flush=True)
def no_output(output):return output
def creates_output(output):output.write_text('TEST',encoding='utf8');return output
recorder_case('19_recorder_no_output',no_output,lambda p,o:{'status':'PASS'})
recorder_case('20_recorder_validation_fails',creates_output,lambda p,o:{'status':'FAIL'})
def drift_output(p,o):p.write_text('DRIFT',encoding='utf8');return {'status':'PASS'}
recorder_case('21_recorder_output_changed',creates_output,drift_output)

f=FIX/'frozen.txt';f.write_text('BEFORE',encoding='utf8');denied=False
try:
    with frozen({str(f):sha(f)}):f.write_text('SHOULD BE DENIED',encoding='utf8')
except PermissionError:denied=True
results.append({'test':'22_frozen_write_denied','result':'PASS' if denied and f.read_text()=='BEFORE' else 'FAIL'})
def additional():
    from uzk_stage5.core import generate,Preview
    from stage7_quality.provenance import verify_receipt
    pre=read(ROOT/'validation/stage7_resume_20261009_175000/ASSEMBLY_01/PREFLIGHT.json')
    base=FIX/'23_accepted_uzk_invalid_preview';base.mkdir()
    rec=Recorder(base,'independent_accepted_uzk',{},{});rec.register('uzk_stage5.core.generate',generate)
    plan=Preview(deepcopy(pre['preview']),'INVALID_PREVIEW')
    try:rec.run('uzk_stage5.core.generate','SAMPLE_ONLY',base/'9_2026 - УЗК (ТЕСТ).docx',
        {'plan':plan,'output_dir':base},lambda p,o:{'status':'PASS'})
    except Exception as e:actual=str(e)
    else:actual='ACCEPTED'
    results.append({'test':'23_accepted_uzk_invalid_preview','result':'PASS' if 'PREVIEW_CHANGED' in actual and not list(base.rglob('STAGE7_OPERATION_RECEIPT.json')) and len(list(base.rglob('FAILURE.json')))==1 else 'FAIL',
        'actual':actual,'accepted_callable':generate.__module__+'.'+generate.__name__,'new_document_created':bool(list(base.glob('*.docx')))})
    r=M['records'][3];receipt=read(r['producer_receipt']['path']);receipt['start']={'path':str(FIX/'missing_START.json'),'sha256':'0'*64}
    p=FIX/'24_fabricated_receipt.json';save(p,receipt)
    try:verify_receipt({'path':str(p),'sha256':sha(p)},M['assembly_run_id'],S.PINS[S.REGISTRY],r['path'],r['sha256'],'SAMPLE_ONLY')
    except Stage7Error as e:actual=str(e)
    else:actual='ACCEPTED'
    results.append({'test':'24_fabricated_success_without_start','result':'PASS' if 'INPUT_OR_CODE_CHANGED' in actual else 'FAIL','actual':actual})
    existing=FIX/'25_recorder_retroactive';existing.mkdir();output=existing/'new.txt';output.write_text('ALREADY GENERATED',encoding='utf8')
    rec=Recorder(existing,'retroactive',{},{});rec.register(creates_output.__module__+'.'+creates_output.__name__,creates_output)
    try:rec.run(creates_output.__module__+'.'+creates_output.__name__,'TEST',output,{'output':output},lambda p,o:{'status':'PASS'})
    except Stage7Error as e:actual=str(e)
    else:actual='ACCEPTED'
    results.append({'test':'25_recorder_existing_output_cannot_retrofit_success','result':'PASS' if 'OUTPUT_EXISTS' in actual and not list(existing.rglob('STAGE7_OPERATION_RECEIPT.json')) and output.read_text()=='ALREADY GENERATED' else 'FAIL','actual':actual})
    save(OUT/'INDEPENDENT_TARGETED_TESTS.json',{'total':len(results),'passed':sum(r['result']=='PASS' for r in results),'tests':results,
        'production':False,'real_source_mutations':False,'full_regression':False,'new_real_quality_generation':False})

additional()
