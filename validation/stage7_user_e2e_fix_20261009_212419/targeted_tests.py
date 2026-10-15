"""Narrow User E2E correction tests; no Stage 1–6 regression or regeneration."""
from copy import deepcopy
from pathlib import Path
import sys
import unittest
from unittest.mock import patch
from uuid import uuid4
sys.dont_write_bytecode=True
HERE=Path(__file__).resolve().parent;ROOT=HERE.parents[1];sys.path.insert(0,str(ROOT))
from stage7_quality.provenance import read,sha,write,verify,digest,Stage7Error,verify_receipt
from stage7_quality.core import verify_manifest,generate_quality,code_bindings
from stage7_quality.word import (parse,table_cells,NS,text,canonical,structural_check,
    references,acts_paragraph,contract_value,validate_quality,STORAGE,BANNER)
from stage7_quality import presentation_revision as revision

SELECTED=read(HERE/'PREFLIGHT.json')['manifest']
MANIFEST=read(SELECTED['path'])
RESULT=read(HERE/'NEW_QUALITY_RESULT.json')
RECEIPT=read(RESULT['receipt']['path'])
OUTPUT=Path(RESULT['output']);DOC=parse(OUTPUT)
FIRST=parse(ROOT.parent/'Гараж казахстан/Паспорт/1_2026 - Документ о качестве.docx')
MANUAL=parse(HERE/'REFERENCE_9_READONLY_SNAPSHOT.docx')
CELLS=table_cells(DOC);REF=table_cells(MANUAL)
FIXTURES=HERE/('isolated_negative_fixtures_'+uuid4().hex[:8]);FIXTURES.mkdir()


def reseal(m):
    m=deepcopy(m);m.pop('seal',None);m['seal']=digest(m);return m


def semantic(n):
    return None if n is None else (n.tag,tuple(sorted(n.attrib.items())),n.text,tuple(semantic(c) for c in n))


class E2ECorrection(unittest.TestCase):
    def reject(self,change):
        m=deepcopy(MANIFEST);change(m)
        with self.assertRaises((Stage7Error,KeyError)):verify_manifest(reseal(m))

    def test_01_point12_exact_xml_and_text(self):
        a=FIRST.findall('w:body/w:tbl',NS)[0].findall('w:tr',NS)[40:49]
        b=DOC.findall('w:body/w:tbl',NS)[0].findall('w:tr',NS)[40:49]
        self.assertEqual([canonical(r) for r in a],[canonical(r) for r in b])
        self.assertIn('от окалины ржавчины по ГОСТ 9.402',text(DOC))

    def test_02_point12_format_mutation_rejected(self):
        bad=deepcopy(DOC);table_cells(bad)[0][42][0].set('SYNTHETIC_format_change','1')
        with self.assertRaises(Stage7Error):structural_check(FIRST,bad)

    def test_03_point13_exact_manual_text(self):
        self.assertEqual(text(CELLS[0][50][0]),text(REF[0][50][0]))
        self.assertEqual(len(CELLS[0][50][0].findall('w:p',NS)),1)

    def test_04_point14_exact_manual_text_and_two_paragraphs(self):
        actual=[text(p) for p in CELLS[0][53][0].findall('w:p',NS)]
        expected=[text(p) for p in REF[0][53][0].findall('w:p',NS)]
        self.assertEqual(actual,expected);self.assertEqual(len(actual),2)
        self.assertEqual(actual[1],STORAGE)

    def test_05_point13_14_four_identical_acts(self):
        self.assertEqual(references(MANIFEST,'basis_point13'),references(MANIFEST,'attachment_point14'))
        self.assertEqual(len(references(MANIFEST,'attachment_point14')),4)

    def test_06_paragraph_and_run_format_matches_manual(self):
        for row in [50,53]:
            a=CELLS[0][row][0].findall('w:p',NS);b=REF[0][row][0].findall('w:p',NS)
            self.assertEqual(len(a),len(b))
            for p,q in zip(a,b):
                self.assertEqual(semantic(p.find('w:pPr',NS)),semantic(q.find('w:pPr',NS)))
                self.assertEqual(semantic(p.find('w:r/w:rPr',NS)),semantic(q.find('w:r/w:rPr',NS)))

    def test_07_dynamic_future_number_and_dates(self):
        m=deepcopy(MANIFEST)
        for r in m['records']:
            r['number']='27/2030-1' if r['document_id']=='vik-1' else '27/2030'
            r['date']='18.07.2030'
        actual=acts_paragraph(m,'basis_point13')
        self.assertIn('№ 27/2030-1 от 18.07.2030 г.',actual)
        self.assertEqual(actual.count('18.07.2030'),4);self.assertNotIn('9/2026',actual)

    def test_08_akz_app_present_but_not_rendered(self):
        self.assertTrue(Path(MANIFEST['records'][5]['path']).is_file())
        for row in [50,53]:self.assertNotIn('Приложение',text(CELLS[0][row][0]))

    def test_09_no_service_status_or_scope_in_points(self):
        points=text(CELLS[0][50][0])+text(CELLS[0][53][0])
        for token in ['ТЕСТ','SAMPLE_ONLY','№0','не подтвержд','3 марки','3 назначенные','Подпись','печать','[',']']:
            self.assertNotIn(token,points)

    def test_10_contract_label_and_value_exact_manual(self):
        for col in [2,3]:self.assertEqual(text(CELLS[0][11][col]),text(REF[0][11][col]))
        self.assertEqual(text(CELLS[0][11][2]),'Договор №')
        self.assertEqual(text(CELLS[0][11][3]).strip(),'06/26 от 03.06.2026')

    def test_11_contract_prefix_only_and_dynamic(self):
        self.assertEqual(contract_value('№ 42/30 от 01.06.2030'),'42/30 от 01.06.2030')
        self.assertEqual(contract_value('42/30 от 01.06.2030'),'42/30 от 01.06.2030')
        with self.assertRaises(Stage7Error):contract_value('№ № 42')

    def test_12_explanation_and_point2_preserved(self):
        self.assertEqual(canonical(CELLS[0][54][0]),canonical(table_cells(FIRST)[0][54][0]))
        self.assertIn('Приложение № 1',text(CELLS[0][15][1]))

    def test_13_test_header_and_metadata_preserved(self):
        self.assertIn(BANNER,text(DOC));self.assertFalse(MANIFEST['official_issue'])
        self.assertTrue(all(r['official_confirmed'] is False for r in MANIFEST['records']))
        self.assertEqual(MANIFEST['records'][3]['status'],'SAMPLE_ONLY')
        self.assertFalse(RECEIPT['results_confirmed']);self.assertFalse(RECEIPT['official_issue'])
        self.assertEqual(MANIFEST['akz_scope']['quantity'],3)

    def test_14_missing_act_rejected(self):
        self.reject(lambda m:m['records'].pop(3))
        self.reject(lambda m:m['records'][1].update(path=str(FIXTURES/'missing_act.docx')))

    def test_15_missing_akz_app_rejected(self):
        self.reject(lambda m:m['records'].pop(5))
        self.reject(lambda m:m['records'][5].update(path=str(FIXTURES/'missing_application.xlsx')))

    def test_16_wrong_status_rejected(self):
        self.reject(lambda m:m['records'][3].update(status='TEST'))
        self.reject(lambda m:m['records'][4].update(status='OFFICIAL'))
        m=deepcopy(MANIFEST);m['records'][3]['status']='OFFICIAL'
        with self.assertRaises(Stage7Error):references(m,'basis_point13')

    def test_17_wrong_date_rejected(self):
        self.reject(lambda m:m['records'][2].update(date='01.01.2020'))

    def test_18_changed_sha_rejected_isolated_copy(self):
        r=MANIFEST['records'][1];p=FIXTURES/'changed_vik.docx'
        with p.open('xb') as f:f.write(Path(r['path']).read_bytes()+b'SYNTHETIC_CHANGED_SHA')
        self.reject(lambda m:m['records'][1].update(path=str(p)))

    def test_19_original_frozen_manifest_and_receipts_unchanged(self):
        verify({SELECTED['path']:SELECTED['sha256']});verify_manifest(MANIFEST)
        verify(read(HERE/'PROTECTED_BEFORE.json'))

    def test_20_exact_code_transition_only(self):
        evidence=revision.verify_producer_code(MANIFEST['code']);self.assertEqual(len(evidence),3)
        bad=dict(MANIFEST['code']);bad[str(ROOT/'stage7_quality/word.py')]='0'*64
        with self.assertRaises(Stage7Error):revision.verify_producer_code(bad)
        bad=dict(MANIFEST['code']);bad[str(ROOT/'passport_app/vom.py')]='0'*64
        with self.assertRaises(Stage7Error):revision.verify_producer_code(bad)

    def test_21_changed_historical_snapshot_rejected(self):
        archive=FIXTURES/'historical_code';archive.mkdir()
        for name in revision.VERSIONS:
            with (archive/name).open('xb') as f:f.write((revision.ARCHIVE/name).read_bytes())
        (archive/'word.py').write_bytes(b'SYNTHETIC_DRIFT')
        with patch.object(revision,'ARCHIVE',archive),self.assertRaises(Stage7Error):
            revision.verify_producer_code(MANIFEST['code'])

    def test_22_native_two_A4_pages_section_tables_points(self):
        native=read(read(RECEIPT['validation']['path'])['checks']['native']['evidence'])
        self.assertEqual((native['pages'],native['sections'],native['tables']),(2,1,2))
        self.assertLess(abs(native['page_width']-595.3),1);self.assertLess(abs(native['page_height']-841.9),1)
        self.assertTrue(native['readonly']);self.assertTrue(native['unchanged_after_close'])
        labels=[r['label'].rstrip('.') for r in native['list_labels'] if r['label'].rstrip('.').isdigit()]
        self.assertEqual(labels,list(map(str,range(1,15))))

    def test_23_new_output_sha_and_receipt_current_code(self):
        self.assertEqual(sha(OUTPUT),RESULT['output_sha256'])
        verify_receipt(RESULT['receipt'],MANIFEST['assembly_run_id'],MANIFEST['membership']['registry_sha256'],str(OUTPUT),sha(OUTPUT),'TEST')
        self.assertEqual(RECEIPT['code'],code_bindings())
        self.assertEqual(RECEIPT['inputs'][SELECTED['path']],SELECTED['sha256'])

    def test_24_no_regeneration_of_other_stages(self):
        self.assertEqual(RESULT['generator_reruns'],{'VIK':0,'UZK':0,'AKZ':0})
        self.assertEqual(RECEIPT['generator'],'stage7_quality.word.write_quality')

    def test_25_output_collision_rejected(self):
        before=sha(OUTPUT)
        with self.assertRaisesRegex(Stage7Error,'OUTPUT_EXISTS'):generate_quality(SELECTED['path'],OUTPUT.parent)
        self.assertEqual(sha(OUTPUT),before)

    def test_26_other_structures_package_and_signatures_preserved(self):
        self.assertTrue(validate_quality(OUTPUT,MANIFEST)['opaque_parts_equal'])
        for i,row in enumerate(table_cells(FIRST)[1]):
            if i!=4:self.assertEqual(canonical(row[0]),canonical(CELLS[1][i][0]))
        self.assertFalse(DOC.xpath('.//w:drawing | .//w:pict',namespaces=NS))

    def test_27_no_false_uzk_reference_when_absent_in_presenter(self):
        m=deepcopy(MANIFEST);m['records']=[r for r in m['records'] if r['type']!='UZK']
        self.assertNotIn('Акт УЗК',acts_paragraph(m,'basis_point13'))
        self.reject(lambda m:m.update(uzk_required=False))


if __name__=='__main__':
    before=read(HERE/'PRODUCT_BEFORE_FINAL_TESTS.json');verify(before)
    with (HERE/'TARGETED_TESTS.log').open('x',encoding='utf8') as log:
        result=unittest.TextTestRunner(stream=log,verbosity=2).run(unittest.defaultTestLoader.loadTestsFromTestCase(E2ECorrection))
    verify(before);write(HERE/'PRODUCT_AFTER_FINAL_TESTS.json',{p:sha(p) for p in before})
    report={'status':'PASS' if result.wasSuccessful() else 'FAIL','tests':result.testsRun,'failures':len(result.failures),
        'errors':len(result.errors),'skipped':len(result.skipped),'failures_detail':[(str(t),error) for t,error in result.failures+result.errors],
        'scope':'Stage7 User E2E correction only; mutation in isolated fixtures/in memory',
        'stage1_6_full_regression':'NOT_RUN','product_before_after_current_equal':True}
    write(HERE/'TARGETED_TEST_RESULTS.json',report);print(report);sys.exit(0 if result.wasSuccessful() else 1)
