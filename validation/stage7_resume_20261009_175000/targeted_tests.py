"""Stage 7 contract tests. Mutations stay in memory or private test fixtures."""
from copy import deepcopy
from pathlib import Path
import json
import sys
import unittest
from uuid import uuid4

ROOT=Path(__file__).resolve().parents[2]
sys.path.insert(0,str(ROOT));sys.dont_write_bytecode=True
from stage7_quality import sources as S
from stage7_quality.core import verify_manifest, generate_quality, code_bindings, validate_uzk
from stage7_quality.provenance import (read,sha,digest,verify,write,Stage7Error,Recorder,
    frozen,verify_receipt,output_boundary,allowed_source)
from stage7_quality.word import (parse,table_cells,text,canonical,NS,structural_check,
    validate_quality,references,BANNER,STORAGE,SCOPE,replace_exact)
from uzk_stage5.core import Preview

HERE=Path(__file__).resolve().parent
ASSEMBLY=HERE/'ASSEMBLY_01'
RESULT=read(ASSEMBLY/'ASSEMBLY_RESULT.json')
MANIFEST=read(RESULT['manifest']['path'])
QUALITY_RECEIPT=read(RESULT['quality_receipt']['path'])
QUALITY=Path(QUALITY_RECEIPT['output'])
PRE=read(ASSEMBLY/'PREFLIGHT.json')
TEMPLATE=parse(S.TEMPLATE);OUTPUT=parse(QUALITY)
FIXTURES=HERE/('failure_fixtures_'+uuid4().hex[:8]);FIXTURES.mkdir()


def reseal(m):
    m=deepcopy(m);m.pop('seal',None);m['seal']=digest(m);return m


class ContractTests(unittest.TestCase):
    def reject(self,change):
        m=deepcopy(MANIFEST);change(m)
        with self.assertRaises((Stage7Error,KeyError)):
            verify_manifest(reseal(m))

    def test_01_exact_template(self):
        self.assertEqual(sha(S.TEMPLATE),S.PINS[S.TEMPLATE])

    def test_02_dynamic_registry_selection(self):
        f=S.current_fields();self.assertEqual(f['number'],'9/2026');self.assertEqual(f['row'],400)

    def test_03_current_date_mass_fields(self):
        f=S.current_fields();self.assertEqual([f[k] for k in ['mass','end_date','vik_date','uzk_date','akz_date']],
            [15822.14,'06.10.2026','29.09.2026','01.10.2026','05.10.2026'])

    def test_04_visible_vom_only(self):
        v=PRE['preview']['vom'];self.assertEqual(len(v['rows']),104);self.assertEqual(len(v['hidden_data_rows']),397)
        self.assertFalse(set(r['row'] for r in v['rows'])&set(v['hidden_data_rows']))
        self.assertEqual(sum(r['quantity'] for r in v['rows']),288);self.assertEqual(v['mass_kg'],'15822.14')

    def test_05_appendix_per_mark(self):
        a=S.validate_appendix(PRE['preview']['vom'],MANIFEST['fields']);self.assertTrue(a['per_mark_equal'])
        self.assertEqual(a['body_rows'],123)
        v=deepcopy(PRE['preview']['vom']);v['rows'][0]['quantity']+=1
        with self.assertRaisesRegex(Stage7Error,'PER_MARK'):S.validate_appendix(v,MANIFEST['fields'])

    def test_06_new_vik_current_dev(self):
        ref=RESULT['vik_receipt'];r=read(ref['path']);self.assertEqual(r['generator'],'passport_app.vik.generate_vik')
        v=read(r['validation']['path'])['checks'];self.assertEqual(v['registry_sha256'],S.PINS[S.REGISTRY])
        actual=read(v['generation_record']);self.assertEqual(actual['registry_sha256'],S.PINS[S.REGISTRY])
        self.assertEqual(actual['output_sha256'],sha(actual['output']))

    def test_07_vik_two_subdocuments(self):
        v=MANIFEST['records'][1:3];self.assertEqual([r['number'] for r in v],['9/2026','9/2026-1'])
        self.assertEqual(v[0]['path'],v[1]['path']);self.assertNotEqual(v[0]['locator'],v[1]['locator'])

    def test_08_old_vik_registry_rejected(self):
        self.reject(lambda m:m['records'][1]['membership'].update(registry_sha256='0'*64))

    def test_09_uzk_receipt_sha_chain(self):
        r=MANIFEST['records'][3]
        receipt=verify_receipt(r['producer_receipt'],MANIFEST['assembly_run_id'],S.PINS[S.REGISTRY],r['path'],r['sha256'],'SAMPLE_ONLY')
        self.assertEqual(receipt['generator'],'uzk_stage5.core.generate')
        self.assertFalse(receipt['results_confirmed'])

    def test_10_uzk_required_present_4_33(self):
        p=Preview(PRE['preview'],PRE['preview_seal']);self.assertEqual(S.validate_uzk_plan(p)['quantity'],33)
        r=MANIFEST['records'][3];self.assertTrue(Path(r['path']).is_file())

    def test_11_required_uzk_missing_rejected(self):
        self.reject(lambda m:m['records'].pop(3))

    def test_12_no_uzk_required_no_false_reference(self):
        # Generic presentation rule, in memory. This approved 9/2026 assembly
        # separately rejects hiding its actual UZK requirement.
        m=deepcopy(MANIFEST);m['records']=[r for r in m['records'] if r['type']!='UZK']
        self.assertFalse(any('УЗК' in s for s in references(m,'basis_point13')))
        self.reject(lambda m:m.update(uzk_required=False))

    def test_13_akz0_pair_integrity(self):
        a=S.validate_akz(PRE['preview']['vom'],MANIFEST['fields']);self.assertEqual(a['pair_id'],'e9fc12e4433feafd')
        self.reject(lambda m:m['records'][5].update(parent_id='other'))

    def test_14_akz_other_variants_rejected(self):
        for n in [1,2,3]:
            with self.subTest(variant=n):self.reject(lambda m:m['records'][4].update(variant=n))

    def test_15_akz_limited_scope_visible(self):
        self.assertEqual(MANIFEST['akz_scope']['quantity'],3)
        self.assertIn('3 назначенные марки / 3 изделия',text(OUTPUT));self.assertIn('не подтверждены',text(OUTPUT))

    def test_16_mixed_runs_rejected(self):
        self.reject(lambda m:m.update(assembly_run_id=uuid4().hex))

    def test_17_explicit_approved_assembly_accepted(self):
        self.assertEqual(verify_manifest(MANIFEST)['passport'],'9/2026')

    def test_18_wrong_document_number_and_date_rejected(self):
        for index in range(6):
            with self.subTest(index=index):
                self.reject(lambda m:m['records'][index].update(number='1/2026'))
                self.reject(lambda m:m['records'][index].update(date='09.09.2026'))

    def test_19_point12_text_equal(self):
        a,b=table_cells(TEMPLATE)[0],table_cells(OUTPUT)[0]
        self.assertEqual([[text(c) for c in row] for row in a[40:49]],[[text(c) for c in row] for row in b[40:49]])

    def test_20_point12_format_equal_and_mutation_rejected(self):
        self.assertTrue(structural_check(TEMPLATE,OUTPUT)['point12_xml_equal'])
        changed=deepcopy(OUTPUT);table_cells(changed)[0][42][0].set('synthetic_format_change','1')
        with self.assertRaises(Stage7Error):structural_check(TEMPLATE,changed)

    def test_21_point13_order_and_content(self):
        self.assertEqual(text(table_cells(OUTPUT)[0][50][0]),'; '.join(references(MANIFEST,'basis_point13'))+'.')
        self.assertEqual(len(references(MANIFEST,'basis_point13')),4)

    def test_22_point14_short_list(self):
        p=table_cells(OUTPUT)[0][53][0].findall('w:p',NS)
        self.assertEqual(len(p),3);self.assertEqual(text(p[1]),STORAGE);self.assertEqual(text(p[2]),SCOPE)
        self.assertEqual(len(references(MANIFEST,'attachment_point14')),5)

    def test_23_no_historical_certificates(self):
        for token in ['АЦСТ -118-03329','АЦСМ-57-00144','114-57813','31889 от']:self.assertNotIn(token,text(OUTPUT))

    def test_24_no_invented_copy_counts(self):
        self.assertTrue(all(r['copies'] is None for r in MANIFEST['records']))
        self.reject(lambda m:m['records'][1].update(copies=2))

    def test_25_missing_file_rejected(self):
        self.reject(lambda m:m['records'][3].update(path=str(FIXTURES/'missing.docx')))

    def test_26_numbering_preserved(self):
        self.assertEqual([canonical(x) for x in TEMPLATE.xpath('.//w:numPr',namespaces=NS)],
                         [canonical(x) for x in OUTPUT.xpath('.//w:numPr',namespaces=NS)])

    def test_27_signatures_unsigned(self):
        self.assertEqual([canonical(c) for ri,row in enumerate(table_cells(TEMPLATE)[1]) if ri!=4 for c in row],
                         [canonical(c) for ri,row in enumerate(table_cells(OUTPUT)[1]) if ri!=4 for c in row])
        self.assertFalse(OUTPUT.xpath('.//w:drawing | .//w:pict',namespaces=NS))

    def test_28_native_word_two_pages(self):
        for name in ['uzk_receipt','vik_receipt','quality_receipt']:
            r=read(RESULT[name]['path']);c=read(r['validation']['path'])['checks']['native'];n=read(c['evidence'])
            self.assertEqual(n['pages'],2);self.assertTrue(n['readonly']);self.assertTrue(n['unchanged_after_close'])
            self.assertEqual(sha(c['evidence']),c['evidence_sha256'])
        n=read(read(QUALITY_RECEIPT['validation']['path'])['checks']['native']['evidence'])
        self.assertEqual((n['sections'],n['tables']),(1,2))
        self.assertEqual([v['label'].rstrip('.') for v in n['list_labels'] if v['label'].rstrip('.').isdigit()],list(map(str,range(1,15))))

    def test_29_output_collision_rejected(self):
        with self.assertRaisesRegex(Stage7Error,'OUTPUT_EXISTS'):output_boundary(QUALITY,QUALITY.parent)
        before=sha(QUALITY)
        with self.assertRaisesRegex(Stage7Error,'OUTPUT_EXISTS'):generate_quality(RESULT['manifest']['path'],QUALITY.parent)
        self.assertEqual(sha(QUALITY),before)

    def test_30_changed_input_after_preview_rejected(self):
        p=FIXTURES/'changed.txt';p.write_text('before');bindings={str(p):sha(p)};p.write_text('after')
        with self.assertRaisesRegex(Stage7Error,'CHANGED'):verify(bindings)
        m=deepcopy(MANIFEST);m['fields']['mass']=0
        with self.assertRaisesRegex(Stage7Error,'MANIFEST_CHANGED'):verify_manifest(m)

    def test_31_production_paths_forbidden(self):
        for name in ['ВЗМК.xlsx','УЗК.xlsm']:
            with self.assertRaisesRegex(Stage7Error,'PRODUCTION'):allowed_source(FIXTURES/name)

    def test_32_stages1_6_unchanged(self):
        verify(read(ROOT/'stage7_quality/accepted_checkpoint.json')['product'])

    def test_33_all_protected_sources_unchanged(self):
        verify(read(HERE/'BASELINE_BEFORE.json'));verify(MANIFEST['inputs'])

    def test_34_test_cannot_become_official(self):
        self.reject(lambda m:m.update(mode='OFFICIAL',official_issue=True))
        self.reject(lambda m:m['records'][4].update(official_confirmed=True))

    def test_35_fabricated_receipt_without_chain_rejected(self):
        fake=FIXTURES/'fabricated_receipt.json'
        write(fake,{'schema':'STAGE7_OPERATION_RECEIPT_V1','state':'COMPLETE_VALIDATED_TEST','status':'SAMPLE_ONLY',
            'assembly_run_id':MANIFEST['assembly_run_id'],'results_confirmed':False,'official_issue':False,
            'output':MANIFEST['records'][3]['path'],'output_sha256':MANIFEST['records'][3]['sha256']})
        r=MANIFEST['records'][3]
        with self.assertRaises((Stage7Error,KeyError)):
            verify_receipt({'path':str(fake),'sha256':sha(fake)},MANIFEST['assembly_run_id'],S.PINS[S.REGISTRY],r['path'],r['sha256'],'SAMPLE_ONLY')

    def test_36_generator_no_output_no_receipt(self):
        def synthetic_no_output():return FIXTURES/'not_created.txt'
        root=FIXTURES/'no_output';root.mkdir();rec=Recorder(root,uuid4().hex,{},{});name=synthetic_no_output.__module__+'.'+synthetic_no_output.__name__
        rec.register(name,synthetic_no_output)
        with self.assertRaisesRegex(Stage7Error,'OUTPUT_MISMATCH'):
            rec.run(name,'TEST',root/'expected.txt',{},lambda p,o:{'status':'PASS'})
        self.assertEqual(len(list(root.rglob('STAGE7_OPERATION_RECEIPT.json'))),0)
        self.assertEqual(len(list(root.rglob('FAILURE.json'))),1)

    def test_37_unexpected_code_drift_rejected_before_call(self):
        p=FIXTURES/'test_code.py';p.write_text('original');expected=sha(p);p.write_text('changed')
        invoked=[]
        def synthetic_never_called():invoked.append(True)
        root=FIXTURES/'code_drift';root.mkdir();rec=Recorder(root,uuid4().hex,{}, {str(p):expected})
        name=synthetic_never_called.__module__+'.'+synthetic_never_called.__name__;rec.register(name,synthetic_never_called)
        with self.assertRaisesRegex(Stage7Error,'CHANGED'):rec.run(name,'TEST',root/'out.txt',{},lambda p,o:{'status':'PASS'})
        self.assertEqual(invoked,[]);self.assertEqual(len(list(root.rglob('STAGE7_OPERATION_RECEIPT.json'))),0)

    def test_38_native_validation_failure_no_receipt(self):
        root=FIXTURES/'validation_failure';root.mkdir();out=root/'out.txt'
        def synthetic_output():out.write_text('SYNTHETIC TEST ONLY');return out
        rec=Recorder(root,uuid4().hex,{},{});name=synthetic_output.__module__+'.'+synthetic_output.__name__;rec.register(name,synthetic_output)
        with self.assertRaisesRegex(Stage7Error,'VALIDATION_FAILED'):rec.run(name,'TEST',out,{},lambda p,o:{'status':'FAIL'})
        self.assertEqual(len(list(root.rglob('RESULT.json'))),1);self.assertEqual(len(list(root.rglob('STAGE7_OPERATION_RECEIPT.json'))),0)

    def test_39_frozen_handle_denies_write_and_delete(self):
        p=FIXTURES/'frozen.txt';p.write_text('original')
        with frozen({str(p):sha(p)}):
            with self.assertRaises(OSError):p.write_text('not allowed')
            with self.assertRaises(OSError):p.unlink()
        self.assertEqual(p.read_text(),'original')

    def test_40_receipt_chain_wrong_operation_rejected(self):
        r=MANIFEST['records'][3]
        with self.assertRaisesRegex(Stage7Error,'IDENTITY'):
            verify_receipt(r['producer_receipt'],uuid4().hex,S.PINS[S.REGISTRY],r['path'],r['sha256'],'SAMPLE_ONLY')

    def test_41_guarded_local_field_rejects_unknown(self):
        p=deepcopy(table_cells(TEMPLATE)[0][16][1])
        with self.assertRaisesRegex(Stage7Error,'AMBIGUOUS'):replace_exact(p,'not a real field','0')

    def test_42_opaque_parts_and_mergefield_preserved(self):
        checks=validate_quality(QUALITY,MANIFEST)
        self.assertTrue(checks['opaque_parts_equal']);self.assertTrue(checks['mergefield_equal'])

    def test_43_order_constant_and_dates(self):
        cells=table_cells(OUTPUT);self.assertEqual(text(cells[0][11][1]),'1')
        self.assertIn('06.10.2026',text(cells[1][4][0]));self.assertIn('15 822,14 кг.',text(cells[0][16][1]))

    def test_44_visible_test_and_scope(self):
        self.assertIn(BANNER,text(OUTPUT));self.assertIn(SCOPE,text(OUTPUT));self.assertFalse(QUALITY_RECEIPT['official_issue'])

    def test_45_uzk_geometry_mutation_rejected(self):
        p=Preview(deepcopy(PRE['preview']),PRE['preview_seal']);p.data['marks'][0]['segments'][0]['L']='999'
        with self.assertRaisesRegex(Stage7Error,'GEOMETRY'):S.validate_uzk_plan(p)

    def test_46_quality_generated_last_actual_times(self):
        q=read(QUALITY_RECEIPT['start']['path'])['time']
        self.assertGreaterEqual(q,read(RESULT['uzk_receipt']['path'])['time'])
        self.assertGreaterEqual(q,read(RESULT['vik_receipt']['path'])['time'])


if __name__=='__main__':
    from datetime import datetime,timezone
    code=code_bindings();write(HERE/'PRODUCT_BEFORE_FINAL_TESTS.json',code)
    stream=(HERE/'TARGETED_TESTS.log').open('x',encoding='utf8')
    suite=unittest.defaultTestLoader.loadTestsFromTestCase(ContractTests)
    result=unittest.TextTestRunner(stream=stream,verbosity=2).run(suite);stream.close()
    verify(code);write(HERE/'PRODUCT_AFTER_FINAL_TESTS.json',{p:sha(p) for p in code})
    report={'time':datetime.now(timezone.utc).isoformat(),'tests':result.testsRun,'failures':len(result.failures),
        'errors':len(result.errors),'skipped':len(result.skipped),'status':'PASS' if result.wasSuccessful() else 'FAIL',
        'failures_detail':[(str(t),detail) for t,detail in result.failures+result.errors],
        'mutation_scope':'In-memory or isolated failure_fixtures only','full_stage1_6_regression':'NOT_RUN',
        'product_hashes_before_after_equal':True}
    write(HERE/'TARGETED_TEST_RESULTS.json',report)
    print(json.dumps(report,ensure_ascii=False,indent=2));sys.exit(0 if result.wasSuccessful() else 1)
