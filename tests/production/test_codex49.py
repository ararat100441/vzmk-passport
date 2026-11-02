"""Compensation boundaries and real Win32 replacement failure recovery."""
from pathlib import Path
from tempfile import TemporaryDirectory
from unittest.mock import patch
from copy import deepcopy
import unittest,hashlib,json,os,subprocess,sys
from production_app import core,compensation as c
from production_app.storage import save,digest,sha,read,Stop

def sealed(path,data):save(path,data|{'seal':digest(data)})

class CompensationTests(unittest.TestCase):
    def setUp(self):
        self.temp=TemporaryDirectory();self.root=Path(self.temp.name);self.id='a'*32
        self.files={r:self.root/(r+suffix) for r,suffix in [('registry','.xlsx'),('catalog','.xlsx'),('uzk','.xlsm')]}
        self.stage=self.root/'.vzmk-operations'/self.id;self.stage.mkdir(parents=True)
        self.journal=self.stage/'journal.json';self.target=self.root/'documents';self.target.mkdir()
        (self.target/'ISSUE.json').write_bytes(b'ISSUE');(self.target/'test.pdf').write_bytes(b'PDF')
        self.plan=self.root/'plan.json';participants=[];manifest_items=[]
        for r,p in self.files.items():
            p.write_bytes(('CURRENT_USER_'+r).encode());backup=self.stage/('backup-'+r+p.suffix);backup.write_bytes(('ORIGINAL_'+r).encode())
            old=self.root/('old_candidate_'+r+p.suffix);old.write_bytes(('OLD_CANDIDATE_'+r).encode())
            after=self.root/('after_'+r+p.suffix);after.write_bytes(('COMPENSATED_USER_'+r).encode())
            before=self.root/('current_copy_'+r+p.suffix);before.write_bytes(p.read_bytes())
            participants.append(dict(role=r,path=str(p),backup=str(backup),candidate=str(old),original_sha256=sha(backup),candidate_sha256=sha(old)))
            manifest_items.append(dict(role=r,path=str(p),before_copy=str(before),candidate=str(after),before_sha256=sha(p),after_sha256=sha(after)))
        sealed(self.plan,dict(schema='VZMK_PRODUCTION_PLAN_1',id=self.id,registry=str(self.files['registry']),target=str(self.target),accounting={r:str(self.files[r]) for r in ['catalog','uzk']}))
        state=dict(id=self.id,phase='NEEDS_READBACK',registry=str(self.files['registry']),plan=str(self.plan),target=str(self.target),participants=participants,output_hashes={'test.pdf':sha(self.target/'test.pdf')},issue_sha256=sha(self.target/'ISSUE.json'))
        core.journal_write(self.journal,state);self.original_journal=self.journal.read_bytes()
        proof=self.root/'proof.json';save(proof,{'status':'PASS_STATIC_CANDIDATES'})
        self.manifest=self.root/'manifest.json'
        sealed(self.manifest,dict(schema=c.SCHEMA,id=self.id,journal=str(self.journal),journal_sha256=sha(self.journal),plan_sha256=sha(self.plan),participants=manifest_items,proofs=[{'path':str(proof),'sha256':sha(proof),'status':'PASS_STATIC_CANDIDATES'}]))
        self.approval={'approved':True,'manifest_sha256':sha(self.manifest),'owner_statement':'EXPLICIT TEST APPROVAL','at':'2026-11-02T00:00:00Z'}
        self.before={r:p.read_bytes() for r,p in self.files.items()}
    def tearDown(self):self.temp.cleanup()
    def assert_before(self):
        self.assertEqual({r:p.read_bytes() for r,p in self.files.items()},self.before)
        self.assertEqual(self.journal.read_bytes(),self.original_journal)
        self.assertTrue((self.target/'test.pdf').is_file())
    def test_success_terminal_receipt_and_no_repeated_recovery(self):
        self.assertEqual(c.apply(self.manifest,self.approval)['phase'],'COMPENSATED')
        self.assertEqual(core.journal_read(self.journal)['phase'],'COMPENSATED')
        self.assertEqual(core.pending(self.files['registry']),[])
        before={r:p.read_bytes() for r,p in self.files.items()}
        self.assertEqual(core.recover(self.journal)['phase'],'COMPENSATED')
        self.assertEqual(before,{r:p.read_bytes() for r,p in self.files.items()})
        self.assertTrue((self.stage/'unpublished_documents/test.pdf').is_file())
        self.assertEqual((self.stage/'compensation/journal.original.json').read_bytes(),self.original_journal)
    def test_closed_receipt_allows_future_live_file_changes(self):
        c.apply(self.manifest,self.approval)
        self.files['registry'].write_bytes(b'NEXT LEGITIMATE ISSUE')
        self.assertEqual(core.pending(self.files['registry']),[])
    def test_no_approval_no_writes(self):
        with self.assertRaises(Stop):c.apply(self.manifest,dict(self.approval,approved=False))
        self.assert_before();self.assertFalse((self.stage/'compensation').exists())
    def test_approval_wrong_manifest_no_writes(self):
        with self.assertRaises(Stop):c.apply(self.manifest,dict(self.approval,manifest_sha256='0'*64))
        self.assert_before()
    def test_new_user_change_blocks_all_writes(self):
        self.files['uzk'].write_bytes(b'NEW USER DATA')
        with self.assertRaises(Stop):c.apply(self.manifest,self.approval)
        self.assertEqual(self.files['registry'].read_bytes(),self.before['registry']);self.assertFalse((self.stage/'compensation').exists())
    def test_candidate_sha_change_blocks(self):
        item=read(self.manifest)['participants'][0];Path(item['candidate']).write_bytes(b'CHANGED')
        with self.assertRaises(Stop):c.apply(self.manifest,self.approval)
        self.assert_before()
    def test_hardlink_before_write_blocks(self):
        alias=self.root/'alias.xlsx';os.link(self.files['registry'],alias)
        try:
            with self.assertRaises(Stop):c.apply(self.manifest,self.approval)
            self.assert_before()
        finally:alias.unlink()
    def test_failure_before_replace_restores_exact_journal(self):
        with patch.object(c,'atomic_replace',side_effect=OSError('injected before replace')):
            with self.assertRaises(Stop):c.apply(self.manifest,self.approval)
        self.assert_before()
    def test_failures_after_each_native_replace_restore_all_CURRENT(self):
        # Each independent fixture exercises real ReplaceFileW, including an
        # error reported after replacement but before the caller records it.
        for at in [1,2,3]:
            with self.subTest(at=at):
                if at>1:self.tearDown();self.setUp()
                real=c.atomic_replace;count=[0]
                def failing(*args,**kwargs):
                    real(*args,**kwargs);count[0]+=1
                    if count[0]==at:raise OSError('injected after native replacement')
                with patch.object(c,'atomic_replace',side_effect=failing):
                    with self.assertRaises(Stop):c.apply(self.manifest,self.approval)
                self.assert_before()
    def test_failure_after_document_move_restores_files_and_folder(self):
        real=c._write_new_json
        def failing(path,data):
            if Path(path).name=='receipt.json':raise OSError('injected receipt write error')
            return real(path,data)
        with patch.object(c,'_write_new_json',side_effect=failing):
            with self.assertRaises(Stop):c.apply(self.manifest,self.approval)
        self.assert_before();self.assertFalse((self.stage/'unpublished_documents').exists())
    def test_partial_native_rename_with_missing_target_restores_CURRENT(self):
        def partial(source,target,*,backup):
            Path(target).rename(backup)
            raise OSError('injected replacement move failure with missing target')
        with patch.object(c,'atomic_replace',side_effect=partial):
            with self.assertRaises(Stop):c.apply(self.manifest,self.approval)
        self.assert_before()
    def test_forged_terminal_phase_without_receipt_blocks(self):
        s=core.journal_read(self.journal);s['phase']='COMPENSATED';core.journal_write(self.journal,s)
        with self.assertRaises((Stop,FileNotFoundError)):core.pending(self.files['registry'])
    def test_tampered_preserved_snapshot_blocks_closed_status(self):
        c.apply(self.manifest,self.approval);(self.stage/'compensation/before-uzk.xlsm').write_bytes(b'CORRUPT')
        with self.assertRaises(Stop):core.pending(self.files['registry'])
    def test_repeat_compensation_blocks(self):
        c.apply(self.manifest,self.approval);before={r:p.read_bytes() for r,p in self.files.items()}
        with self.assertRaises(Stop):c.apply(self.manifest,self.approval)
        self.assertEqual(before,{r:p.read_bytes() for r,p in self.files.items()})
    def test_audit_tamper_blocks_closed_status(self):
        c.apply(self.manifest,self.approval);audit=self.stage/'compensation/audit.jsonl';audit.write_bytes(audit.read_bytes().replace(b'INTENT_VERIFIED',b'INTENT_TAMPERED'))
        with self.assertRaises(Stop):core.pending(self.files['registry'])
    def test_failure_after_terminal_write_restores_CURRENT_and_original_journal(self):
        real=core.journal_write
        def failing(path,state):
            real(path,state)
            if state['phase']=='COMPENSATED':raise OSError('injected after final journal write')
        with patch.object(core,'journal_write',side_effect=failing):
            with self.assertRaises(Stop):c.apply(self.manifest,self.approval)
        self.assert_before()
    def test_process_crash_intent_recovers_to_CURRENT(self):
        source=Path(c.__file__).parents[1]
        script="import sys,os;sys.path.insert(0,sys.argv[1]);from production_app import compensation as c;from production_app.storage import read;real=c.atomic_replace\ndef crash(*a,**k):\n real(*a,**k);os._exit(91)\nc.atomic_replace=crash;c.apply(sys.argv[2],read(sys.argv[3]))"
        approval=self.root/'approval.json';save(approval,self.approval)
        child=subprocess.run([sys.executable,'-c',script,str(source),str(self.manifest),str(approval)],creationflags=getattr(subprocess,'CREATE_NO_WINDOW',0))
        self.assertEqual(child.returncode,91)
        self.assertEqual(read(self.journal)['phase'],'COMPENSATING')
        result=core.recover(self.journal);self.assertEqual(result['compensation'],'FAILED_RESTORED_CURRENT');self.assert_before()

if __name__=='__main__':unittest.main()
