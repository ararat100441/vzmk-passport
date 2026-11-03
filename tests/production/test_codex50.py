"""Real temporary journal hardlinks; Office mocks isolate the path regression.

Native Office and frozen GUI acceptance are separate mandatory real-copy gates.
"""
import os,copy,unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch,MagicMock
import test_codex46 as base
from production_app import core,storage


class JournalAfterWrite(base.PostCommitSafety):
    def linked(self):
        alias=self.root/'journal-temporary-second-name.json'
        os.link(self.journal,alias)
        self.addCleanup(lambda:alias.unlink(missing_ok=True))
        with self.assertRaisesRegex(storage.Stop,'Hardlink'):storage.safe(self.journal)
        return alias

    def prepared_readback(self):
        self.state['backup']=self.state['participants'][0]['backup']
        self.plan.update(project='TEST',number='11/2026')
        self.plan.pop('seal');self.plan['seal']=storage.digest(self.plan)
        storage.save(self.run/'plan.json',self.plan)
        for n in ['prepared.json','verified.json','confirmation.json','release.json']:
            value=storage.read(self.run/n);value['plan_seal']=self.plan['seal'];storage.save(self.run/n,value)
        conf=storage.read(self.run/'confirmation.json');conf['verified_sha256']=storage.sha(self.run/'verified.json');storage.save(self.run/'confirmation.json',conf)
        self.pdf=self.target/(self.target.name+'.pdf');self.pdf.write_bytes(b'isolated PDF test boundary')
        self.state['output_hashes']={self.pdf.name:storage.sha(self.pdf)};self.write_state()

    def test_linked_journal_readback_calls_native_verification(self):
        self.prepared_readback();self.linked()
        registry=SimpleNamespace(projects={'TEST':[dict(number='11/2026',row=2)]},sheet=1,columns={'number':1})
        app=MagicMock();app.Workbooks.Open.return_value.Worksheets.return_value.Cells.return_value.Value2='11/2026'
        with patch('production_app.coordinated.native_readback') as native, patch('production_app.core.excel_session') as excel, patch('production_app.core.read_registry',return_value=registry), patch('production_app.print_pdf.verify_package_print',return_value={'filename':self.pdf.name,'pages':1,'sha256':storage.sha(self.pdf)}), patch('production_app.print_pdf.verify_print_pdf'), patch('production_app.core.office.verify_documents'):
            excel.return_value.__enter__.return_value=app
            self.assertEqual(core.readback(self.journal)['status'],'PASS')
            native.assert_called_once();app.Workbooks.Open.assert_called_once()

    def test_linked_journal_complete_atomic_update_and_terminal_pending(self):
        self.prepared_readback();alias=self.linked();before=alias.read_bytes()
        with patch('production_app.core.supervisor_proof'):
            result=core.complete(self.journal,self.root/'readback-result.json')
        self.assertEqual(result['status'],'PASS');self.assertEqual(core.journal_read(self.journal)['phase'],'COMPLETE')
        self.assertEqual(alias.read_bytes(),before)
        self.assertEqual(core.pending(self.registry),[])
        self.assertTrue(storage.safe(self.journal))

    def test_linked_corrupted_journal_stops_before_office(self):
        self.linked();self.journal.write_text('{invalid',encoding='utf8')
        with patch('production_app.coordinated.native_readback') as native:
            with self.assertRaises(Exception):core.readback(self.journal)
            native.assert_not_called()

    def test_linked_wrong_seal_stops(self):
        alias=self.linked();record=storage.read(self.journal);record['id']='b'*32
        self.journal.write_text(__import__('json').dumps(record),encoding='utf8')
        with self.assertRaisesRegex(storage.Stop,'Журнал повреждён'):core.readback(self.journal)

    def test_linked_wrong_participant_sha_stops(self):
        self.linked();self.catalog.write_bytes(b'user changed')
        with patch('production_app.coordinated.native_readback') as native:
            with self.assertRaisesRegex(storage.Stop,'После записи изменён'):core.readback(self.journal)
            native.assert_not_called()

    def test_linked_resealed_participant_substitution_stops(self):
        foreign=self.root/'foreign.xlsx';foreign.write_bytes(self.catalog.read_bytes())
        self.state['participants'][1]['path']=str(foreign);self.write_state();self.linked()
        with self.assertRaisesRegex(storage.Stop,'Участник'):core.readback(self.journal)

    def test_linked_wrong_phase_still_stops(self):
        self.state['phase']='STARTED';self.write_state();self.linked()
        with self.assertRaisesRegex(storage.Stop,'фаза'):core.readback(self.journal)

    def test_linked_relocated_journal_stops(self):
        self.linked();foreign=self.root/'journal.json';os.link(self.journal,foreign)
        with self.assertRaisesRegex(storage.Stop,'Журнал вне'):core.readback(foreign)

    def test_linked_journal_reparse_metadata_stops(self):
        self.linked();actual=storage.windows_file_identity
        def metadata(p):
            result=actual(p)
            if Path(p)==self.journal:result['attributes']|=0x400
            return result
        with patch('production_app.storage.windows_file_identity',side_effect=metadata):
            with self.assertRaisesRegex(storage.Stop,'reparse'):core.readback(self.journal)

    def test_strict_recovery_does_not_inherit_exception(self):
        self.linked()
        with self.assertRaisesRegex(storage.Stop,'Hardlink'):core.recover(self.journal)

    def test_linked_complete_rejects_changed_participant(self):
        self.prepared_readback();self.linked();self.catalog.write_bytes(b'changed after readback')
        with patch('production_app.core.supervisor_proof'):
            with self.assertRaisesRegex(storage.Stop,'После readback изменён'):core.complete(self.journal,self.root/'result.json')
        self.assertEqual(core.journal_read(self.journal)['phase'],'NEEDS_READBACK')

    def test_linked_resealed_wrong_operation_id_stops(self):
        self.state['id']='b'*32;self.write_state();self.linked()
        with self.assertRaisesRegex(storage.Stop,'Пути транзакции'):core.readback(self.journal)


if __name__=='__main__':unittest.main()
