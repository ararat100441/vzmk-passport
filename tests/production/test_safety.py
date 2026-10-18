"""Safety/refusal and native Windows recovery tests on disposable files."""
import copy,json,os,tempfile,unittest,subprocess
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch
import sys
ROOT=Path(__file__).resolve().parents[2];sys.path.insert(0,str(ROOT))
from production_app import model,core,facts
from production_app.storage import Stop,sha,save,read,mutex
from passport_app.excel import locked_file

class Safety(unittest.TestCase):
    def setUp(self):
        self.temp=tempfile.TemporaryDirectory();self.root=Path(self.temp.name)
    def tearDown(self):self.temp.cleanup()
    def test_number_series_preserved(self):
        for value in ('1/2026','1/1/2026','19/2/2027'):
            n,s,y=model.parse_number(value);self.assertEqual(model.number(n,s,y),value)
        for value in ('1-1/2026','0/1/2026','1/0/2026','1/1/26','1/2026/1','1/1/2026/2'):
            with self.subTest(value=value),self.assertRaises(Stop):model.parse_number(value)
    def test_mixed_series_stop_before_read_or_write(self):
        r=SimpleNamespace(projects={'ТЭМПО 1':[{'number':'1/1/2026','row':342},{'number':'10/2026','row':401}]})
        with patch('production_app.model.load_book') as load:
            with self.assertRaisesRegex(Stop,'КОНФЛИКТ СЕРИИ'):model.select(r,'ТЭМПО 1')
            load.assert_not_called()
    def tempo_registry_fixture(self):
        from datetime import datetime
        rows=[342,351,355,366,370,378,380,388,394,401,407]
        nums=[f'{n}/1/2026' for n in range(1,10)]+['10/2026','11/2026']
        entries=[{'row':row,'number':value,'date':datetime(2026,9,12),
                  'object':'ТЭМПО 1 объект'}
                 for row,value in zip(rows,nums)]
        registry=SimpleNamespace(projects={'ТЭМПО 1':entries},path=self.root/'registry.xlsx',
                                 sheet='Паспорт качества')
        fields={r:{1:'МК-3465 от 05.05.2026',2:'АО КЗМК ТЭМПО',
                    3:'ТЭМПО 1 объект',4:'SKL-VD-(L3-41-030)-N-ZX 2511.3.135.021'}
                for r in rows}
        class Sheet:
            def cell(self,row,col):return SimpleNamespace(value=fields[row][col])
        class Book:
            def __getitem__(self,name):return Sheet()
            def close(self):pass
        return registry,fields,Book()

    def test_tempo1_approved_legacy_aliases_compute_12_without_editing_sources(self):
        registry,fields,book=self.tempo_registry_fixture()
        before=copy.deepcopy(registry.projects)
        with patch('production_app.model.columns',return_value={'CONTRACT':1,'CUSTOMER':2,'OBJECT':3,'CIPHER':4}),patch('production_app.model.load_book',return_value=book):
            p=model.select(registry,'ТЭМПО 1')
        self.assertEqual(p.first_number,'1/1/2026')
        self.assertEqual(p.last_number,'11/2026')  # historical source remains literal
        self.assertEqual(p.next_number,'12/1/2026')
        self.assertEqual(registry.projects,before)
        self.assertEqual([x['number'] for x in registry.projects['ТЭМПО 1'][-2:]],['10/2026','11/2026'])

    def test_tempo1_legacy_aliases_refuse_dev_kazakhstan_requisites(self):
        registry,fields,book=self.tempo_registry_fixture()
        fields[401][1]='06/26 от 03.06.2026'
        fields[401][4]='WAA0007144-AC2-КМД'
        with patch('production_app.model.columns',return_value={'CONTRACT':1,'CUSTOMER':2,'OBJECT':3,'CIPHER':4}),patch('production_app.model.load_book',return_value=book):
            with self.assertRaisesRegex(Stop,'КОНФЛИКТ РЕКВИЗИТОВ'):model.select(registry,'ТЭМПО 1')

    def test_tempo1_extra_mixed_series_does_not_bypass_stop(self):
        registry,fields,book=self.tempo_registry_fixture()
        registry.projects['ТЭМПО 1'].append({'row':411,'number':'12/2026'})
        with patch('production_app.model.load_book') as load:
            with self.assertRaisesRegex(Stop,'КОНФЛИКТ СЕРИИ'):model.select(registry,'ТЭМПО 1')
            load.assert_not_called()

    def test_tempo1_duplicate_normalized_number_is_refused(self):
        registry,fields,book=self.tempo_registry_fixture()
        registry.projects['ТЭМПО 1'].append({'row':409,'number':'10/1/2026'})
        with patch('production_app.model.load_book') as load:
            with self.assertRaisesRegex(Stop,'Дубли номеров с учётом'):model.select(registry,'ТЭМПО 1')
            load.assert_not_called()

    def test_tempo1_legacy_aliases_require_exact_rows_and_first(self):
        registry,fields,book=self.tempo_registry_fixture()
        registry.projects['ТЭМПО 1'][0]['row']=340
        with patch('production_app.model.load_book') as load:
            with self.assertRaisesRegex(Stop,'не совпали со строками'):model.select(registry,'ТЭМПО 1')
            load.assert_not_called()

    def test_tempo1_one_corrected_number_is_still_safe(self):
        registry,fields,book=self.tempo_registry_fixture()
        registry.projects['ТЭМПО 1'][-1]['number']='11/1/2026'
        with patch('production_app.model.columns',return_value={'CONTRACT':1,'CUSTOMER':2,'OBJECT':3,'CIPHER':4}),patch('production_app.model.load_book',return_value=book):
            p=model.select(registry,'ТЭМПО 1')
        self.assertEqual(p.next_number,'12/1/2026')

    def test_tempo2_mixed_series_is_not_authorized(self):
        registry,fields,book=self.tempo_registry_fixture()
        registry.projects={'ТЭМПО 2':registry.projects['ТЭМПО 1']}
        with patch('production_app.model.load_book') as load:
            with self.assertRaisesRegex(Stop,'КОНФЛИКТ СЕРИИ'):model.select(registry,'ТЭМПО 2')
            load.assert_not_called()

    def test_training_cannot_target_real_registry(self):
        registry=self.root/'ВЗМК.xlsx';registry.write_bytes(b'original')
        with self.assertRaises((Stop,FileNotFoundError)):core.training_guard({'training':True,'registry':str(registry),'output':str(self.root)})
        self.assertEqual(registry.read_bytes(),b'original')
    def valid_facts(self):
        evidence=self.root/'protocol.txt';evidence.write_text('Подтверждено ответственным: фактический контроль выполнен.',encoding='utf8')
        f={'project':'P','number':'1/1/2026','kind':'ACTUAL','confirmed':True,'authorized':True,'forms_approved':True,'responsible':'Иванов',
            'start_date':'2026-10-01','date':'2026-10-18','vik_date':'2026-10-11','uzk_date':'2026-10-13','akz_date':'2026-10-17','evidence':[{'path':str(evidence),'sha256':sha(evidence)}],
            'marks':[{'mark':'M1','profile':'Профиль','paint':'','akz':None,'uzk_quantity':1,'uzk_result':'ГОДЕН','basis':'Покрытие не требуется по проекту'}],'climates':[]}
        v=SimpleNamespace(rows=(('M1','Балка',2,1,2,1,2),))
        return f,v
    def test_actual_requires_explicit_facts(self):
        f,v=self.valid_facts();self.assertEqual(facts.validate(f,v,'P','1/1/2026'),[])
        self.assertIsNone(f['marks'][0]['akz'])
        for key in ('confirmed','authorized','forms_approved','responsible'):
            bad=copy.deepcopy(f);bad[key]=False
            with self.subTest(key=key),self.assertRaises(Stop):facts.validate(bad,v,'P','1/1/2026')
    def test_test_results_cannot_be_actual(self):
        f,v=self.valid_facts()
        for kind in ('TRAINING','TEST','SAMPLE_ONLY'):
            bad=copy.deepcopy(f);bad['kind']=kind
            with self.subTest(kind=kind),self.assertRaises(Stop):facts.validate(bad,v,'P','1/1/2026')
        path=Path(f['evidence'][0]['path']);path.write_text('SAMPLE_ONLY TEST — НЕ ЯВЛЯЕТСЯ КОНТРОЛЕМ',encoding='utf8');f['evidence'][0]['sha256']=sha(path)
        with self.assertRaisesRegex(Stop,'модельные'):facts.validate(f,v,'P','1/1/2026')
    def test_no_automatic_akz_zero(self):
        f,v=self.valid_facts();f['marks'][0]['basis']=''
        with self.assertRaisesRegex(Stop,'отсутствие АКЗ'):facts.validate(f,v,'P','1/1/2026')
        self.assertIsNone(f['marks'][0]['akz'])
    def test_missing_mark_or_excess_uzk_stop(self):
        f,v=self.valid_facts();f['marks'][0]['uzk_quantity']=3
        with self.assertRaises(Stop):facts.validate(f,v,'P','1/1/2026')
        f['marks']=[]
        with self.assertRaises(Stop):facts.validate(f,v,'P','1/1/2026')
    def test_unconfirmed_climate_stop(self):
        f,v=self.valid_facts();f['marks'][0].update(akz=0,paint='Подтверждённый материал')
        with self.assertRaises(Stop):facts.validate(f,v,'P','1/1/2026')
        f['climates']=[dict(variant=0,air=20,surface=12,humidity=50,dew=11,thickness=100,material='M',basis='Протокол')]
        with self.assertRaisesRegex(Stop,'точки росы'):facts.validate(f,v,'P','1/1/2026')
    def test_office_proof_required(self):
        p=self.root/'fake.json';save(p,{'status':'PASS','task':'release','plan':'P'})
        with self.assertRaises((Stop,KeyError)):core.supervisor_proof(p,'release','P')
    def crash_state(self,phase='REPLACE_INTENT',after_replace=True):
        # Arbitrary bytes intentionally isolate NTFS recovery from document logic.
        registry=self.root/'COPY_ВЗМК.xlsx';registry.write_bytes(b'ORIGINAL_REGISTRY_BYTES')
        token='1'*32;stage=self.root/'.vzmk-operations'/token;stage.mkdir(parents=True)
        backup=stage/'backup.xlsx';backup.write_bytes(registry.read_bytes());before=sha(registry)
        target=self.root/'new_documents';target.mkdir();doc=target/'output.docx';doc.write_bytes(b'PREPARED_DOCUMENT')
        save(target/'ISSUE.json',{'training':True})
        new=stage/'candidate.xlsx';new.write_bytes(b'STAGED_REGISTRY_BYTES');candidate=sha(new)
        if after_replace:
            from passport_app.excel import atomic_replace
            atomic_replace(new,registry)
        state={'id':token,'phase':phase,'registry':str(registry),'backup':str(backup),'target':str(target),'original_sha256':before,'candidate_sha256':candidate,
            'new':True,'output_hashes':{'output.docx':sha(doc)},'issue_sha256':sha(target/'ISSUE.json')}
        journal=stage/'journal.json';core.journal_write(journal,state)
        return journal,registry,target,before
    def test_recovery_after_replace_before_journal_update(self):
        journal,registry,target,before=self.crash_state()
        result=core.recover(journal);self.assertEqual(result['phase'],'ROLLED_BACK')
        self.assertEqual(sha(registry),before);self.assertFalse(target.exists())
        self.assertTrue((journal.parent/'unpublished_documents/output.docx').is_file())
        self.assertEqual(core.recover(journal)['phase'],'ROLLED_BACK')
    def test_recovery_before_replace(self):
        journal,registry,target,before=self.crash_state(phase='DOCUMENTS_PUBLISHED',after_replace=False)
        core.recover(journal);self.assertEqual(sha(registry),before);self.assertFalse(target.exists())
    def test_recovery_never_overwrites_parallel_changes(self):
        journal,registry,target,before=self.crash_state();registry.write_bytes(b'USER_EDIT')
        with self.assertRaisesRegex(Stop,'другим пользователем'):core.recover(journal)
        self.assertEqual(registry.read_bytes(),b'USER_EDIT');self.assertTrue(target.exists())
    def test_recovery_never_moves_edited_documents(self):
        journal,registry,target,before=self.crash_state();(target/'output.docx').write_bytes(b'USER_EDIT')
        h=sha(registry)
        with self.assertRaises(Stop):core.recover(journal)
        self.assertEqual(sha(registry),h);self.assertTrue(target.exists())
    def test_corrupt_backup_and_journal_refused(self):
        journal,registry,target,before=self.crash_state();(journal.parent/'backup.xlsx').write_bytes(b'corrupt')
        with self.assertRaises(Stop):core.recover(journal)
        data=read(journal);data['phase']='COMPLETE';save(journal,data)
        with self.assertRaisesRegex(Stop,'Журнал повреждён'):core.pending(registry)
    def test_native_exclusive_lock_refuses_second_writer(self):
        p=self.root/'registry.xlsx';p.write_bytes(b'protected')
        with locked_file(p,writable=True):
            with self.assertRaises(Exception):
                with locked_file(p,writable=True):self.fail('lock not exclusive')
        self.assertEqual(p.read_bytes(),b'protected')
    def test_other_process_mutex_refuses_parallel_operation(self):
        p=self.root/'registry.xlsx';p.write_bytes(b'protected')
        script="import sys; from production_app.storage import mutex;\nwith mutex(sys.argv[1]):\n print('LOCKED',flush=True)\n sys.stdin.read(1)"
        proc=subprocess.Popen([sys.executable,'-c',script,str(p)],cwd=ROOT,stdin=subprocess.PIPE,stdout=subprocess.PIPE,stderr=subprocess.PIPE,text=True)
        try:
            self.assertEqual(proc.stdout.readline().strip(),'LOCKED')
            with self.assertRaisesRegex(Stop,'другим экземпляром'):
                with mutex(p):self.fail('parallel mutex acquired')
        finally:proc.communicate('x',timeout=10)
        self.assertEqual(proc.returncode,0)

if __name__=='__main__':unittest.main(verbosity=2)
