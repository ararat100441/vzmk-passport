"""The approved assembly is explicit and fixed. Quality is always generated last."""
from copy import deepcopy
from pathlib import Path
from uuid import uuid4
from .provenance import (Stage7Error, Recorder, sha, read, write, verify, frozen,
                         digest, now, verify_receipt, output_boundary)
from . import sources as S


def code_bindings():
    checkpoint = read(Path(__file__).with_name('accepted_checkpoint.json'))
    bindings = checkpoint['product']
    verify(bindings)
    new = [*Path(__file__).parent.glob('*.py'), Path(__file__).with_name('render_emf.ps1'),
           Path(__file__).with_name('accepted_checkpoint.json'), S.ROOT/'stage7.py',
           S.ROOT/'Запустить Этап 7 DEV.cmd']
    return bindings | {str(p.resolve()): sha(p) for p in new}


def preflight():
    from uzk_stage5.core import preview
    from uzk_stage5.result_samples import load_sample_results
    verify({str(p): h for p,h in S.PINS.items()})
    code = code_bindings()
    fields = S.current_fields()
    plan = preview(S.PROJECT, S.NUMBER, S.VOM)
    uzk = S.validate_uzk_plan(plan)
    vom = plan.data['vom']
    if len(vom['rows']) != 104 or len(vom['hidden_data_rows']) != 397 or sum(r['quantity'] for r in vom['rows']) != 288:
        raise Stage7Error('VISIBLE_VOM_CONTRACT_MISMATCH')
    appendix = S.validate_appendix(vom,fields)
    akz = S.validate_akz(vom,fields)
    sample = load_sample_results(plan.data)
    checkpoint = read(Path(__file__).with_name('accepted_checkpoint.json'))
    inputs = {str(p): h for p,h in S.PINS.items()} | plan.data['fingerprints'] | sample['fingerprints']
    inputs |= read(S.AKZ_ROOT/'current_preview.json')['payload']['sources']
    inputs |= checkpoint['extra_inputs']
    inputs |= {str(S.PAIR):sha(S.PAIR),str(S.AKZ_ROOT/'current_preview.json'):sha(S.AKZ_ROOT/'current_preview.json')}
    verify(inputs | code)
    return {'fields':fields,'plan':plan,'code':code,'inputs':inputs,'appendix':appendix,
            'akz':akz,'uzk':uzk,'sample':sample}


def validate_uzk(path, operation, plan):
    from uzk_stage5.result_samples import extract_rows, TEST_LABEL, UNCONFIRMED_LABEL, cell_text
    from .word import parse, NS
    from .native import inspect
    rows = extract_rows(Path(path).read_bytes(),path)
    marks = plan.data['marks']
    if len(rows)!=4 or [r['mark'] for r in rows]!=[m['mark'] for m in marks] or any((r['graph7'],r['graph8'])!=('ДНО','годен') for r in rows):
        raise Stage7Error('UZK_OUTPUT_MARKS_OR_SAMPLE_MISMATCH')
    root=parse(path)
    table=root.findall('w:body/w:tbl',NS)[0]
    body=table.findall('w:tr',NS)[5:]
    def fmt(v):
        from decimal import Decimal
        return format(Decimal(str(v)).normalize(),'f').replace('.',',')
    for row,m in zip(body,marks):
        cells=row.findall('w:tc',NS)
        expected={1:'\n'.join('L='+fmt(s['L'])+'х'+fmt(s['S'])+' мм' for s in m['segments']),
                  3:'01.10.2026',4:'\n'.join('5' for s in m['segments']),5:'\n'.join('70' for s in m['segments'])}
        # Stage 5's geometry fmt uses a decimal point; frequency uses a comma.
        expected[1]=expected[1].replace(',','.')
        if any(cell_text(cells[i])!=v for i,v in expected.items()):
            raise Stage7Error('UZK_OUTPUT_GEOMETRY_MISMATCH')
    native=inspect(path,operation,kind='UZK',required=[TEST_LABEL,UNCONFIRMED_LABEL,'9/2026','01.10.2026'])
    return {'status':'PASS','native':native,'geometry_verified':True,'marks':4,'controlled_units':33,
            'sample_only':True,'actual_control_confirmed':False,'preview_seal':plan.seal}


def validate_vik(path, operation, snapshot_root, fields):
    from .native import inspect
    evidence=list((snapshot_root/'.vik-stage').glob('*/generation.json'))
    if len(evidence)!=1:
        raise Stage7Error('VIK_ACTUAL_RUN_RECORD_MISSING')
    record=read(evidence[0])
    if (record['registry_sha256']!=S.PINS[S.REGISTRY] or record['output_sha256']!=sha(path)
            or record['template_sha256']!=sha(S.VIK_TEMPLATE) or record['row']!=fields['row']
            or not record['normal_word_reopen'] or not record['constant_text_preserved']):
        raise Stage7Error('VIK_CURRENT_REGISTRY_OR_GENERATION_MISMATCH')
    native=inspect(path,operation,kind='VIK',required=['9/2026','9/2026-1','29.09.2026',fields['object'],fields['cipher'],fields['connections']])
    pages=read(native['evidence'])['page_texts']
    if '9/2026' not in pages[0] or '9/2026-1' not in pages[1] or any('29.09.2026' not in p for p in pages):
        raise Stage7Error('VIK_SUBDOCUMENTS_MISMATCH')
    return {'status':'PASS','native':native,'generation_record':str(evidence[0]),
            'generation_sha256':sha(evidence[0]),'registry_sha256':record['registry_sha256'],
            'two_current_acts':True,'connections':fields['connections']}


def membership(fields):
    return {'project':S.PROJECT,'passport':S.NUMBER,'registry_path':str(S.REGISTRY),
        'registry_sha256':S.PINS[S.REGISTRY],'registry_row':fields['row'],
        'vom_path':str(S.VOM),'vom_sha256':S.PINS[S.VOM],
        'visible_rows':104,'hidden_rows_excluded':397,'quantity':288,'mass_kg':'15822.14'}


def make_manifest(run_id, pre, uzk_receipt, vik_receipt):
    member=membership(pre['fields'])
    v=read(vik_receipt['path']);u=read(uzk_receipt['path'])
    common={'membership':member,'copies':None,'official_confirmed':False}
    def record(identity,kind,number,day,path,status,title,usage,**extra):
        return common | {'document_id':identity,'type':kind,'number':number,'date':day,
            'path':str(path),'sha256':sha(path),'status':status,'display_title':title,
            'usage':usage,'pair_id':None,'parent_id':None,'variant':None,'source_run_id':None,
            'producer_receipt':None} | extra
    records=[record('appendix1','QUALITY_APPENDIX_1',S.NUMBER,'06.10.2026',S.APPENDIX,'VERIFIED_TEST',
        'Приложение № 1 к документу о качестве № 9/2026 от 06.10.2026 г.', ['construction_point2'],
        verification=pre['appendix'],source_run_id=None,provenance_note='Existing reviewed file; original generator run unknown; not attributed to Stage 7'),
        record('vik','VIK',S.NUMBER,'29.09.2026',v['output'],'TEST',
            'Акт ВИК № 9/2026 от 29.09.2026 г. [ТЕСТ]', ['basis_point13','attachment_point14'],
            source_run_id=v['operation_id'],producer_receipt=vik_receipt,locator='Native Word page 1'),
        record('vik-1','VIK',S.NUMBER+'-1','29.09.2026',v['output'],'TEST',
            'Акт ВИК № 9/2026-1 от 29.09.2026 г. [ТЕСТ]', ['basis_point13','attachment_point14'],
            source_run_id=v['operation_id'],producer_receipt=vik_receipt,locator='Native Word page 2'),
        record('uzk','UZK',S.NUMBER,'01.10.2026',u['output'],'SAMPLE_ONLY',
            'Акт УЗК № 9/2026 от 01.10.2026 г. [SAMPLE_ONLY: образец, контроль не подтверждён]', ['basis_point13','attachment_point14'],
            source_run_id=u['operation_id'],producer_receipt=uzk_receipt,verification=pre['uzk']),
        record('akz0','AKZ',S.NUMBER,'05.10.2026',S.AKZ,'TEST',
            'Сертификат качества антикоррозийной обработки № 9/2026 от 05.10.2026 г. [ТЕСТ, АКЗ №0]', ['basis_point13','attachment_point14'],
            source_run_id=pre['akz']['pair_id'],pair_id=pre['akz']['pair_id'],variant=0,
            producer_receipt={'path':str(S.PAIR),'sha256':sha(S.PAIR)},verification=pre['akz']),
        record('akz0-appendix','AKZ_APPENDIX',S.NUMBER,'05.10.2026',S.AKZ_APP,'TEST',
            'Приложение № 9/2026 к акту АКЗ № 9/2026 от 05.10.2026 г. [ТЕСТ, АКЗ №0]', ['attachment_point14'],
            source_run_id=pre['akz']['pair_id'],pair_id=pre['akz']['pair_id'],parent_id='akz0',variant=0,
            producer_receipt={'path':str(S.PAIR),'sha256':sha(S.PAIR)})]
    m={'schema':'STAGE7_SELECTED_TEST_MANIFEST_V1','assembly_run_id':run_id,'created_at':now(),
        'mode':'DEV_TEST_ONLY','state':'FROZEN_VERIFIED_TEST_ASSEMBLY','official_issue':False,
        'project':S.PROJECT,'passport':S.NUMBER,'membership':member,'fields':pre['fields'],
        'uzk_required':True,'akz_scope':pre['akz'],'inputs':pre['inputs'],'code':pre['code'],
        'records':records,'user_authorization':'Explicit Stage 7 9/2026 assembly: current Appendix1, new VIK, new SAMPLE_ONLY UZK, accepted AKZ0 pair only'}
    return m | {'seal':digest(m)}


def verify_manifest(manifest):
    m=deepcopy(manifest);seal=m.pop('seal',None)
    if seal!=digest(m):
        raise Stage7Error('MANIFEST_CHANGED_REPEAT_PREVIEW')
    if (m['schema']!='STAGE7_SELECTED_TEST_MANIFEST_V1' or m['mode']!='DEV_TEST_ONLY'
            or m['state']!='FROZEN_VERIFIED_TEST_ASSEMBLY' or m['official_issue'] is not False
            or m['project']!=S.PROJECT or m['passport']!=S.NUMBER or m['uzk_required'] is not True):
        raise Stage7Error('MANIFEST_SCOPE_OR_OFFICIAL_FORBIDDEN')
    verify(m['inputs'])
    from .presentation_revision import verify_producer_code
    verify_producer_code(m['code'])
    fields=S.current_fields()
    if fields!=m['fields'] or m['membership']!=membership(fields):
        raise Stage7Error('MANIFEST_CURRENT_MEMBERSHIP_MISMATCH')
    expected=[('appendix1','QUALITY_APPENDIX_1',S.NUMBER,'06.10.2026','VERIFIED_TEST'),
        ('vik','VIK',S.NUMBER,'29.09.2026','TEST'),('vik-1','VIK',S.NUMBER+'-1','29.09.2026','TEST'),
        ('uzk','UZK',S.NUMBER,'01.10.2026','SAMPLE_ONLY'),('akz0','AKZ',S.NUMBER,'05.10.2026','TEST'),
        ('akz0-appendix','AKZ_APPENDIX',S.NUMBER,'05.10.2026','TEST')]
    records=m['records']
    if [(r['document_id'],r['type'],r['number'],r['date'],r['status']) for r in records]!=expected:
        raise Stage7Error('SELECTED_RECORDS_MISMATCH_OR_MISSING')
    for r in records:
        if r['membership']!=m['membership'] or r['official_confirmed'] is not False or r['copies'] is not None:
            raise Stage7Error('RECORD_MEMBERSHIP_OR_OFFICIAL_MISMATCH')
        verify({r['path']:r['sha256']})
    for r,p in [(records[0],S.APPENDIX),(records[4],S.AKZ),(records[5],S.AKZ_APP)]:
        if r['path']!=str(p) or r['sha256']!=S.PINS[p]:
            raise Stage7Error('REVIEWED_SOURCE_SUBSTITUTED')
    for r,generator in [(records[1],'passport_app.vik.generate_vik'),(records[2],'passport_app.vik.generate_vik'),(records[3],'uzk_stage5.core.generate')]:
        receipt=verify_receipt(r['producer_receipt'],m['assembly_run_id'],S.PINS[S.REGISTRY],r['path'],r['sha256'],r['status'])
        if receipt['generator']!=generator or receipt['operation_id']!=r['source_run_id']:
            raise Stage7Error('WRONG_GENERATOR_OR_RUN')
    if records[1]['path']!=records[2]['path'] or records[1]['source_run_id']!=records[2]['source_run_id']:
        raise Stage7Error('VIK_TWO_ACTS_DIFFERENT_RUNS')
    if (any(r['variant']!=0 for r in records[4:]) or records[4]['pair_id']!=records[5]['pair_id']
            or records[5]['parent_id']!='akz0' or records[4]['pair_id']!=read(S.PAIR)['act_id']):
        raise Stage7Error('AKZ_PAIR_OR_VARIANT_MISMATCH')
    expected_titles={
        'vik':'Акт ВИК № 9/2026 от 29.09.2026 г. [ТЕСТ]',
        'vik-1':'Акт ВИК № 9/2026-1 от 29.09.2026 г. [ТЕСТ]',
        'uzk':'Акт УЗК № 9/2026 от 01.10.2026 г. [SAMPLE_ONLY: образец, контроль не подтверждён]',
        'akz0':'Сертификат качества антикоррозийной обработки № 9/2026 от 05.10.2026 г. [ТЕСТ, АКЗ №0]',
        'akz0-appendix':'Приложение № 9/2026 к акту АКЗ № 9/2026 от 05.10.2026 г. [ТЕСТ, АКЗ №0]'}
    for i,r in enumerate(records):
        expected_usage=['construction_point2'] if i==0 else ['attachment_point14'] if i==5 else ['basis_point13','attachment_point14']
        if r['usage']!=expected_usage or i>0 and r['display_title']!=expected_titles[r['document_id']]:
            raise Stage7Error('RECORD_PRESENTATION_MISMATCH')
    return m


def generate_quality(manifest_path, output_dir):
    from .word import write_quality, validate_quality, BANNER
    from .native import inspect
    manifest_path=Path(manifest_path).resolve()
    manifest=read(manifest_path)
    verify_manifest(manifest)
    output_dir=Path(output_dir).resolve()
    allowed=[S.ROOT/'validation',S.ROOT/'DEV/Stage7_Выход']
    if not any(output_dir.is_relative_to(p) and output_dir!=p for p in allowed):
        raise Stage7Error('QUALITY_OUTPUT_PATH_FORBIDDEN')
    output=output_boundary(output_dir/'9_2026 - Документ о качестве (ТЕСТ).docx',output_dir)
    inputs=manifest['inputs'] | {str(manifest_path):sha(manifest_path)}
    for r in manifest['records']:
        inputs[r['path']]=r['sha256']
        if r.get('producer_receipt'):
            inputs[r['producer_receipt']['path']]=r['producer_receipt']['sha256']
    from .presentation_revision import historical_code_evidence
    inputs |= historical_code_evidence(manifest['code'])
    recorder=Recorder(output_dir,manifest['assembly_run_id'],inputs,code_bindings())
    recorder.register('stage7_quality.word.write_quality',write_quality)
    def validate(path,operation):
        verify_manifest(manifest)
        xml=validate_quality(path,manifest)
        fields=manifest['fields']
        mass=format(float(fields['mass']),',.2f').replace(',',' ').replace('.',',')+' кг.'
        native=inspect(path,operation,kind='QUALITY',required=[BANNER,fields['number'],fields['end_date'],mass])
        return {'status':'PASS','xml':xml,'native':native}
    return recorder.run('stage7_quality.word.write_quality','TEST',output,
        {'manifest':manifest,'output':output},validate)


def assemble(run_root):
    from uzk_stage5.core import generate
    from passport_app.vik import generate_vik
    from passport_app.config import Settings
    run_root=Path(run_root).resolve()
    if not run_root.is_relative_to(S.ROOT/'validation') or run_root==S.ROOT/'validation':
        raise Stage7Error('ASSEMBLY_VALIDATION_ONLY')
    run_root.mkdir(parents=True,exist_ok=False)
    run_id=uuid4().hex
    pre=preflight()
    write(run_root/'PREFLIGHT.json',{k:v for k,v in pre.items() if k!='plan'} | {'preview':pre['plan'].data,'preview_seal':pre['plan'].seal})
    recorder=Recorder(run_root,run_id,pre['inputs'],pre['code'])
    recorder.register('uzk_stage5.core.generate',generate)
    uzk_dir=run_root/'NEW_SAMPLE_ONLY_UZK'
    # START is written before the accepted generator is invoked.
    uzk=recorder.run('uzk_stage5.core.generate','SAMPLE_ONLY',uzk_dir/'9_2026 - УЗК (ТЕСТ).docx',
        {'plan':pre['plan'],'output_dir':uzk_dir},lambda p,o:validate_uzk(p,o,pre['plan']))
    snapshot=run_root/'VIK_ISOLATION'
    passport=snapshot/'documents/Гараж казахстан/Паспорт'
    with frozen(pre['inputs'] | pre['code']):
        (passport/'9-2026 от 06.10.2026').mkdir(parents=True,exist_ok=False)
        (snapshot/'DEV_registry.xlsx').write_bytes(S.REGISTRY.read_bytes())
        (passport/'1_2026 - ВИК.docx').write_bytes(S.VIK_TEMPLATE.read_bytes())
    settings=Settings(snapshot/'documents',snapshot/'DEV_registry.xlsx',
        {S.PROJECT:{'passport_root':'Гараж казахстан/Паспорт'}})
    recorder.register('passport_app.vik.generate_vik',generate_vik)
    vik=recorder.run('passport_app.vik.generate_vik','TEST',passport/'9-2026 от 06.10.2026/9_2026 - ВИК.docx',
        {'settings':settings,'project_name':S.PROJECT,'number':S.NUMBER,'test_root':snapshot},
        lambda p,o:validate_vik(p,o,snapshot,pre['fields']))
    manifest=make_manifest(run_id,pre,uzk,vik)
    verify_manifest(manifest)
    selected=write(run_root/'STAGE7_SELECTED_TEST_MANIFEST.json',manifest)
    quality=generate_quality(selected['path'],run_root/'QUALITY_TEST')
    result={'status':'COMPLETE_VALIDATED_TEST','assembly_run_id':run_id,'manifest':selected,
            'uzk_receipt':uzk,'vik_receipt':vik,'quality_receipt':quality,'official_issue':False}
    write(run_root/'ASSEMBLY_RESULT.json',result)
    return result
