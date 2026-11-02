"""The post-commit exception is confined to sealed participant paths/SHA.

Native Office and full frozen GUI acceptance run separately on real copies.
"""
import copy,os,tempfile,unittest
from pathlib import Path
from unittest.mock import patch
from production_app import core,coordinated,storage,shared_uzk


class PostCommitSafety(unittest.TestCase):
    def setUp(self):
        self.temp=tempfile.TemporaryDirectory();self.addCleanup(self.temp.cleanup)
        self.root=Path(self.temp.name).resolve();self.local=self.root/'profile'
        self.id='a'*32;self.run=self.local/'jobs'/self.id;self.run.mkdir(parents=True)
        self.registry=self.root/'registry.xlsx';self.catalog=self.root/'catalog.xlsx';self.uzk=self.root/'uzk.xlsm'
        self.target=self.root/'published';self.target.mkdir()
        self.stage=self.root/'.vzmk-operations'/self.id;self.stage.mkdir(parents=True)
        self.journal=self.stage/'journal.json';items=[]
        for role,path in [('registry',self.registry),('catalog',self.catalog),('uzk',self.uzk)]:
            path.write_bytes(('candidate-'+role).encode())
            backup=self.stage/('backup-'+role+path.suffix);backup.write_bytes(('original-'+role).encode())
            candidate=self.run/(role+path.suffix);candidate.write_bytes(path.read_bytes())
            items.append(dict(role=role,path=str(path),backup=str(backup),candidate=str(candidate),
                              original_sha256=storage.sha(backup),candidate_sha256=storage.sha(candidate)))
        plan=dict(schema='VZMK_PRODUCTION_PLAN_1',id=self.id,run=str(self.run),registry=str(self.registry),
                  target=str(self.target),cfg={'registry':str(self.registry)},facts={'simple_ui':True},
                  accounting={'catalog':str(self.catalog),'uzk':str(self.uzk)},
                  fingerprints={i['path']:i['original_sha256'] for i in items})
        plan['seal']=storage.digest(plan);self.plan=plan;storage.save(self.run/'plan.json',plan)
        storage.save(self.run/'prepared.json',dict(plan_seal=plan['seal'],registry=items[0]['candidate'],
                     snapshot_sha256=items[0]['candidate_sha256'],accounting=items[1:]))
        storage.save(self.run/'verified.json',dict(plan_seal=plan['seal']))
        storage.save(self.run/'confirmation.json',dict(plan_seal=plan['seal'],reviewed=True,approved=True,
                     verified_sha256=storage.sha(self.run/'verified.json')))
        storage.save(self.run/'release.json',dict(plan_seal=plan['seal']))
        storage.save(self.target/'ISSUE.json',{'test':True})
        self.state=dict(id=self.id,phase='NEEDS_READBACK',plan=str(self.run/'plan.json'),registry=str(self.registry),
                        target=str(self.target),participants=items,new=True,original_sha256=items[0]['original_sha256'],
                        candidate_sha256=items[0]['candidate_sha256'],output_hashes={},issue_sha256=storage.sha(self.target/'ISSUE.json'))
        self.write_state()
        self.local_patch=patch('production_app.core.local_root',return_value=self.local);self.local_patch.start();self.addCleanup(self.local_patch.stop)

    def write_state(self):core.journal_write(self.journal,self.state)
    def context(self):return core.post_commit_context(self.journal)

    def test_default_rejects_real_hardlink_before_write(self):
        alias=self.root/'other.xlsx';os.link(self.registry,alias)
        for path in (self.registry,alias):
            with self.subTest(path=path),self.assertRaisesRegex(storage.Stop,'Hardlink'):storage.safe(path)
        with self.assertRaisesRegex(storage.Stop,'Hardlink'):
            with storage.mutex(self.registry):pass
        with self.assertRaisesRegex(storage.Stop,'Hardlink'):shared_uzk.read_shared_uzk(self.registry)

    def test_normal_post_commit_paths_and_candidate_hashes(self):
        before={i['path']:storage.sha(i['path']) for i in self.state['participants']}
        s,d,r,t=self.context();self.assertTrue(coordinated.verify_current(s))
        self.assertEqual(r,self.registry);self.assertEqual(t,self.target)
        self.assertEqual(before,{p:storage.sha(p) for p in before})

    def test_real_post_commit_hardlinks_do_not_repeat_count_gate(self):
        for i in self.state['participants']:os.link(i['path'],self.root/('alias-'+Path(i['path']).name))
        for _ in range(2):
            s,_,_,_=self.context();self.assertTrue(coordinated.verify_current(s))
            with storage.mutex(self.registry,post_commit=True):pass
        with self.assertRaisesRegex(storage.Stop,'Hardlink'):storage.safe(self.registry)

    def test_wrong_candidate_sha_stops_before_office(self):
        self.catalog.write_bytes(b'foreign bytes')
        with patch('production_app.coordinated.native_readback') as office:
            with self.assertRaisesRegex(storage.Stop,'После записи изменён'):core.readback(self.journal)
            office.assert_not_called()

    def test_reparse_remains_blocked_after_commit(self):
        actual=storage.windows_file_identity
        def reparse(path):
            value=actual(path)
            if Path(path)==self.registry:value['attributes']|=0x400
            return value
        with patch('production_app.storage.windows_file_identity',side_effect=reparse):
            with self.assertRaisesRegex(storage.Stop,'reparse'):self.context()

    def test_native_metadata_failure_is_not_suppressed(self):
        with patch('production_app.storage.windows_file_identity',side_effect=storage.Stop('native query failed')):
            with self.assertRaisesRegex(storage.Stop,'native query failed'):storage.post_commit_path(self.registry)

    def test_wrong_phase_stops(self):
        self.state['phase']='STARTED';self.write_state()
        with self.assertRaisesRegex(storage.Stop,'фаза'):self.context()

    def test_journal_seal_corruption_stops(self):
        record=storage.read(self.journal);record['phase']='COMPLETE';storage.save(self.journal,record)
        with self.assertRaisesRegex(storage.Stop,'Журнал повреждён'):self.context()

    def test_plan_seal_corruption_stops(self):
        plan=copy.deepcopy(self.plan);plan['target']=str(self.root/'foreign');storage.save(self.run/'plan.json',plan)
        with self.assertRaisesRegex(storage.Stop,'План повреждён'):self.context()

    def test_even_resealed_participant_path_substitution_stops(self):
        foreign=self.root/'same-bytes.xlsx';foreign.write_bytes(self.catalog.read_bytes())
        self.state['participants'][1]['path']=str(foreign);self.write_state()
        with self.assertRaisesRegex(storage.Stop,'Участник'):self.context()

    def test_even_resealed_duplicate_role_stops(self):
        self.state['participants'][1]['role']='uzk';self.write_state()
        with self.assertRaisesRegex(storage.Stop,'участники'):self.context()

    def test_wrong_candidate_binding_stops(self):
        self.state['participants'][1]['candidate_sha256']='0'*64;self.write_state()
        with self.assertRaisesRegex(storage.Stop,'Участник'):self.context()

    def test_unconfirmed_transaction_stops(self):
        conf=storage.read(self.run/'confirmation.json');conf['approved']=False;storage.save(self.run/'confirmation.json',conf)
        with self.assertRaisesRegex(storage.Stop,'подтверждения'):self.context()

    def test_corrupt_backup_stops(self):
        Path(self.state['participants'][1]['backup']).write_bytes(b'foreign backup')
        with self.assertRaisesRegex(storage.Stop,'Резервная копия'):self.context()

    def test_changed_issue_and_escaped_document_paths_stop(self):
        storage.save(self.target/'ISSUE.json',{'changed':True})
        with self.assertRaisesRegex(storage.Stop,'сведения'):self.context()
        self.state['issue_sha256']=storage.sha(self.target/'ISSUE.json');self.state['output_hashes']={'../foreign.pdf':'0'*64};self.write_state()
        with self.assertRaisesRegex(storage.Stop,'путь документа'):self.context()


if __name__=='__main__':unittest.main()
