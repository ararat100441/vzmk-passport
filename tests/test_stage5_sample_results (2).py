"""Focused sample-results checks against fresh output and immutable old evidence."""
from copy import deepcopy
import json
import os
from pathlib import Path
import sys
import unittest
from unittest.mock import patch
from zipfile import ZipFile

from docx import Document
from docx.table import _Cell
from lxml import etree as E

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
import uzk_stage5.core as C
import uzk_stage5.result_samples as S
from uzk_stage5.inputs import Stage5Error, sha, registry_rows, mark_profiles
from uzk_stage5.word import NS, contents, write_test_document
from profiles_stage4.storage import load_catalog
from profiles_stage4.offline import no_network


def read(path):
    return json.loads(Path(path).read_text(encoding='utf-8-sig'))


class SampleResultsTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.run_dir = Path(os.environ['STAGE5_SAMPLE_VALIDATION_DIR']).resolve()
        assert cls.run_dir.is_relative_to(ROOT/'validation')
        cls.baseline = read(cls.run_dir/'baseline.json')
        cls.actual = read(cls.run_dir/'current_preview.json')
        cls.data = cls.actual['preview']
        cls.new = read(cls.run_dir/'new_document.json')
        cls.output = Path(cls.new['output'])
        cls.historical = read(cls.run_dir/'historical_rows.json')
        cls.native = read(cls.run_dir/'native_word.json')
        with ZipFile(cls.output) as archive:
            cls.doc = E.fromstring(archive.read('word/document.xml'))
        cls.rows = cls.doc.find('w:body', NS).findall('w:tbl', NS)[0].findall('w:tr', NS)[5:]

    def test_01_independent_exact_extraction_all_56_rows(self):
        independent = []
        for source in read(self.run_dir/'historical_sources.json'):
            path = Path(self.data['project_root'])/source['relative_path']
            self.assertEqual(sha(path), source['sha256'])
            document = Document(path)
            table = document.tables[0]
            headers=[]
            for index,row in enumerate(table._tbl.tr_lst):
                if len(row.tc_lst)==8:
                    values=[_Cell(tc,table).text for tc in row.tc_lst]
                    if 'Обнаруженные дефекты' in values[6].replace('\n',' ') and 'Заключение' in values[7]:
                        headers.append(index)
            self.assertEqual(len(headers),1)
            rows = table._tbl.tr_lst[headers[0]+1:]
            self.assertEqual(len(rows), source['rows'])
            for row in rows:
                cells = [_Cell(tc, table).text for tc in row.tc_lst]
                independent.append((path.name, cells[0], cells[6], cells[7]))
        expected = [(r['act'], r['mark'], r['graph7'], r['graph8']) for r in self.historical]
        self.assertEqual(independent, expected)
        self.assertEqual(len(independent), 56)
        self.assertEqual({(r[2], r[3]) for r in independent}, {('ДНО','годен')})

    def test_02_accepted_sample_is_unconfirmed(self):
        sample = S.load_sample_results(self.data)
        self.assertEqual((sample['graph7'],sample['graph8']), ('ДНО','годен'))
        self.assertEqual(sample['row_count'],56)
        self.assertEqual(len(sample['fingerprints']),10)
        self.assertEqual(sample['mode'],'SAMPLE_ONLY')
        self.assertIs(sample['confirmed'],False)

    def test_03_empty_missing_or_blank_sample_stops(self):
        for rows in ([], [{'mark':'A','graph7':'','graph8':'годен'}],
                     [{'mark':'A','graph7':'ДНО','graph8':' '}],
                     [{'mark':'','graph7':'ДНО','graph8':'годен'}]):
            with self.subTest(rows=rows), self.assertRaisesRegex(Stage5Error,'SAMPLE_NOT_PROVEN'):
                S.assess_rows(rows)

    def test_04_incompatible_variants_stop_without_normalization(self):
        for alternative in ('Годен','не годен','годен '):
            rows=[{'mark':'A','graph7':'ДНО','graph8':'годен'},
                  {'mark':'B','graph7':'ДНО','graph8':alternative}]
            with self.subTest(alternative=alternative), self.assertRaisesRegex(Stage5Error,'SAMPLE_AMBIGUOUS'):
                S.assess_rows(rows)

    def test_05_missing_source_stops_before_document_creation(self):
        data=deepcopy(self.data);data['project_root']=str(self.run_dir/'missing_project')
        output=self.run_dir/'missing_sample_must_not_exist.docx'
        with self.assertRaisesRegex(Stage5Error,'SAMPLE_SOURCE_MISSING'):
            write_test_document(data,output)
        self.assertFalse(output.exists())

    def test_06_source_drift_stops_readonly(self):
        sources=list(S.ACCEPTED_SOURCES)
        name,_,count=sources[0];sources[0]=(name,'0'*64,count)
        with patch.object(S,'ACCEPTED_SOURCES',tuple(sources)):
            with self.assertRaisesRegex(Stage5Error,'SAMPLE_SOURCE_CHANGED'):
                S.load_sample_results(self.data)
        sample=S.load_sample_results(self.data)
        sample['fingerprints'][str(self.run_dir/'absent_source')]='0'*64
        with self.assertRaisesRegex(Stage5Error,'SAMPLE_SOURCE_CHANGED'):
            S.verify_sample_sources(sample)

    def test_07_other_project_and_official_issuance_blocked(self):
        data=deepcopy(self.data);data['passport']['project']='Other'
        with self.assertRaisesRegex(Stage5Error,'SAMPLE_PROJECT_NOT_PROVEN'):
            S.load_sample_results(data)
        for mode in ('OFFICIAL','CONFIRMED',None):
            data=deepcopy(self.data);data['kind']=mode
            data['results_confirmed']=True
            data['responsible_executor']='FIXTURE'
            with self.subTest(mode=mode), self.assertRaisesRegex(Stage5Error,'OFFICIAL_RESULTS_NOT_IMPLEMENTED'):
                S.load_sample_results(data)

    def test_08_all_four_rows_have_exact_sample_text(self):
        self.assertEqual(len(self.rows),4)
        self.assertEqual([contents(r.findall('w:tc',NS)[0]) for r in self.rows],
                         ['Бп2-1','Пс1-9','Пс1-10','Пс2-16'])
        for row in self.rows:
            cells=row.findall('w:tc',NS)
            self.assertEqual(len(cells),8)
            self.assertEqual((contents(cells[6]),contents(cells[7])),('ДНО','годен'))

    def test_09_explicit_visible_test_and_unconfirmed_labels(self):
        body=self.doc.find('w:body',NS)
        titles=[p for p in body.findall('w:p',NS) if contents(p).startswith('Акт №')]
        self.assertEqual(len(titles),1)
        text=contents(titles[0])
        self.assertIn('Акт № 9/2026 от 01.10.2026',text)
        self.assertIn(S.TEST_LABEL,text)
        self.assertIn(S.UNCONFIRMED_LABEL,text)
        self.assertTrue(titles[0].findall('.//w:br',NS))
        for run in titles[0].findall('w:r',NS):
            self.assertIsNone(run.find('w:rPr/w:vanish',NS))

    def test_10_only_authorized_document_changes(self):
        old=self.baseline['existing_test_docx']['path']
        with ZipFile(old) as a, ZipFile(self.output) as b:
            self.assertEqual(a.namelist(),b.namelist())
            self.assertEqual([n for n in a.namelist() if a.read(n)!=b.read(n)],['word/document.xml'])
            previous=E.fromstring(a.read('word/document.xml'))
        current=deepcopy(self.doc)
        oldbody=previous.find('w:body',NS);body=current.find('w:body',NS)
        oldtitle=next(p for p in oldbody.findall('w:p',NS) if contents(p).startswith('Акт №'))
        title=next(p for p in body.findall('w:p',NS) if contents(p).startswith('Акт №'))
        body.replace(title,deepcopy(oldtitle))
        oldrows=oldbody.findall('w:tbl',NS)[0].findall('w:tr',NS)[5:]
        rows=body.findall('w:tbl',NS)[0].findall('w:tr',NS)[5:]
        for oldrow,row in zip(oldrows,rows):
            oldcells=oldrow.findall('w:tc',NS);cells=row.findall('w:tc',NS)
            for i in range(6):
                self.assertEqual(E.tostring(oldcells[i]),E.tostring(cells[i]))
            for i in (6,7):
                # The cell/paragraph/run formatting remains the retained template's.
                for prop in ('w:tcPr','w:p/w:pPr','w:p/w:r/w:rPr'):
                    a,b=oldcells[i].find(prop,NS),cells[i].find(prop,NS)
                    self.assertEqual(E.tostring(a) if a is not None else None,
                                     E.tostring(b) if b is not None else None)
                row.replace(cells[i],deepcopy(oldcells[i]))
        self.assertEqual(E.tostring(previous),E.tostring(current))

    def test_11_actual_calculation_and_geometry_unchanged(self):
        old=read(ROOT/'validation/stage5_implementation_20261008_105100/final_run.json')['preview']
        self.assertEqual(self.data['marks'],old['marks'])
        stock,_=registry_rows()
        mapping=mark_profiles(self.data['mark_catalog'],required_marks=[m['mark'] for m in self.data['marks']])
        actual,_=C.calculate_marks('Гараж Казахстан',self.data['vom']['rows'],stock,mapping,load_catalog())
        self.assertEqual(actual,self.data['marks'])
        expected={'Бп2-1':(8,[('1280',5),('1024',8.4)]),'Пс1-9':(9,[('3600',4)]),
                  'Пс1-10':(15,[('6000',4)]),'Пс2-16':(1,[('480',3)])}
        self.assertEqual(sum(m['Q'] for m in actual),33)
        for mark in actual:
            q,segments=expected[mark['mark']]
            self.assertEqual(mark['Q'],q)
            self.assertEqual([(s['L'],s['S']) for s in mark['segments']],segments)
            self.assertTrue(all(s['frequency_mhz']==5 and s['angle_degrees']==70 for s in mark['segments']))

    def test_12_s5r1_core_and_inputs_unchanged(self):
        for name in ('uzk_stage5/core.py','uzk_stage5/inputs.py'):
            path=ROOT/name
            self.assertEqual(sha(path),self.baseline['product_before'][str(path)])

    def test_13_sources_old_evidence_and_old_test_unchanged(self):
        allowed={str(ROOT/p) for p in ('uzk_stage5/word.py','uzk_stage5/gui.py','STAGE5_DEV.md')}
        changes=[p for p,h in self.baseline['protected_before'].items() if sha(p)!=h]
        self.assertEqual(set(changes),allowed)
        old=self.baseline['existing_test_docx']
        self.assertEqual(sha(old['path']),old['sha256'])

    def test_14_native_word_normal_open_and_pages(self):
        n=self.native
        self.assertEqual(n['status'],'PASS_NORMAL_READONLY_OPEN')
        self.assertTrue(n['readonly']);self.assertTrue(n['unchanged'])
        self.assertEqual(n['before_sha'],self.new['sha256'])
        self.assertEqual((n['pages'],n['sections'],n['tables'],n['first_table_rows'],n['hyperlinks']),
                         (2,1,2,9,0))
        self.assertTrue(n['owned_instance'])
        self.assertEqual(len(n['images']),2)

    def test_15_no_stale_hyperlinks(self):
        self.assertFalse(self.doc.findall('.//w:hyperlink',NS))
        with ZipFile(self.output) as archive:
            rels=E.fromstring(archive.read('word/_rels/document.xml.rels'))
        self.assertFalse(any(r.get('Type','').endswith('/hyperlink') for r in rels))

    def test_16_raw_extractor_preserves_literal_characters(self):
        cell=E.fromstring(('<w:tc xmlns:w="'+NS['w']+'"><w:p><w:r><w:t> ДНО </w:t>'
            '<w:tab/><w:t>годен</w:t><w:br/><w:t>Х</w:t></w:r></w:p>'
            '<w:p><w:r><w:t>строка</w:t></w:r></w:p></w:tc>').encode())
        self.assertEqual(S.cell_text(cell),' ДНО \tгоден\nХ\nстрока')


if __name__=='__main__':
    run=Path(os.environ['STAGE5_SAMPLE_VALIDATION_DIR']).resolve()
    access=[]
    def guard(event,args):
        if event!='open' or not isinstance(args[0],(str,bytes)):
            return
        path=Path(os.fsdecode(args[0])).resolve()
        if path.name.casefold() in ('взмк.xlsx','узк.xlsm'):
            raise AssertionError('PRODUCTION_ACCESS_FORBIDDEN: '+str(path))
        mode,flags=args[1:3]
        writing=(isinstance(mode,str) and any(c in mode for c in 'wax+')) or (
            isinstance(flags,int) and flags & (os.O_WRONLY|os.O_RDWR|os.O_CREAT|os.O_TRUNC))
        if writing and not path.is_relative_to(run):
            raise AssertionError('WRITE_OUTSIDE_VALIDATION: '+str(path))
        access.append({'path':str(path),'writing':bool(writing)})
    sys.addaudithook(guard)
    paths=[*C.product_paths(),ROOT/'STAGE5_DEV.md',ROOT/'tests/test_stage5.py',
           ROOT/'tests/test_stage5_s5_r1.py',Path(__file__).resolve()]
    before={str(p):sha(p) for p in paths}
    (run/'product_before_tests.json').write_text(json.dumps(before,ensure_ascii=False,indent=2),encoding='utf8')
    with no_network() as attempts, (run/'targeted_tests.log').open('w',encoding='utf8') as log:
        result=unittest.TextTestRunner(stream=log,verbosity=2).run(unittest.defaultTestLoader.loadTestsFromTestCase(SampleResultsTests))
    after={str(p):sha(p) for p in paths}
    assert before==after,'Tested product changed'
    (run/'product_after_tests.json').write_text(json.dumps(after,ensure_ascii=False,indent=2),encoding='utf8')
    summary={'tests_run':result.testsRun,'successful':result.wasSuccessful(),'network_attempts':attempts,
             'failures':[(str(t),s) for t,s in result.failures],'errors':[(str(t),s) for t,s in result.errors]}
    (run/'targeted_tests.json').write_text(json.dumps(summary,ensure_ascii=False,indent=2),encoding='utf8')
    (run/'targeted_access_guard.json').write_text(json.dumps({'events':access,'production_reads':0,
        'writes_outside_validation':0},ensure_ascii=False,indent=2),encoding='utf8')
    print(json.dumps(summary,ensure_ascii=False,indent=2))
    raise SystemExit(0 if result.wasSuccessful() else 1)
