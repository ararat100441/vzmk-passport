"""Native Windows coverage for optional archive and preservation on failure."""
from datetime import datetime
from pathlib import Path
import subprocess
import sys
from tempfile import TemporaryDirectory
import unittest
from unittest.mock import patch

from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table
from passport_app.registry import REQUIRED
from production_app import catalog_ledger, coordinated, core, model, service_archive as archive
from production_app.storage import digest, read, save, sha, Stop


@unittest.skipUnless(sys.platform == 'win32', 'Windows rename and locking')
class ArchiveTests(unittest.TestCase):
    def setUp(self):
        tmp = TemporaryDirectory()
        self.addCleanup(tmp.cleanup)
        self.root = Path(tmp.name)
        self.id = 'd'*32
        self.stage = self.root/'.vzmk-operations'/self.id
        self.stage.mkdir(parents=True)
        self.journal = self.stage/'journal.json'
        self.plan = self.root/'jobs'/self.id/'plan.json'
        self.target = self.root/'documents'
        self.target.mkdir()
        items = []
        for role, folder, ext in [('registry', self.root, '.xlsx'),
                                  ('catalog', self.root/'project', '.xlsx'),
                                  ('uzk', self.root, '.xlsm')]:
            folder.mkdir(exist_ok=True)
            p = folder/(role + ext)
            p.write_bytes(b'updated ' + role.encode())
            b = self.stage/('backup-' + role + ext)
            b.write_bytes(b'original ' + role.encode())
            archive.replacement_path(p, self.id).write_bytes(b.read_bytes())
            items.append({'role': role, 'path': str(p), 'backup': str(b),
                          'original_sha256': sha(b), 'candidate_sha256': sha(p)})
        save(self.plan, {'id': self.id})
        save(self.target/'ISSUE.json', {'files': {}})
        pdf = self.target/'documents.pdf'
        pdf.write_bytes(b'fixture PDF')
        self.state = {'id': self.id, 'phase': 'NEEDS_READBACK', 'participants': items,
                      'registry': items[0]['path'], 'target': str(self.target), 'plan': str(self.plan),
                      'backup': items[0]['backup'], 'issue_sha256': sha(self.target/'ISSUE.json'),
                      'output_hashes': {pdf.name: sha(pdf)}}
        core.journal_write(self.journal, self.state)

    def source(self, i):
        return archive.replacement_path(self.state['participants'][i]['path'], self.id)

    def test_no_archive_before_complete_then_idempotent(self):
        with self.assertRaisesRegex(Stop, 'COMPLETE'):
            archive.archive_completed(self.journal)
        self.assertTrue(all(self.source(i).is_file() for i in range(3)))
        moved = coordinated.complete(self.journal, self.state)['service_archive']
        self.assertEqual(len(moved['moved']), 3)
        self.assertEqual(moved['retained'], [])
        terminal = sha(self.journal)
        self.assertEqual(archive.archive_completed(self.journal), moved)
        self.assertEqual(sha(self.journal), terminal)
        for rec in moved['moved']:
            self.assertFalse(Path(rec['source']).exists())
            self.assertEqual(sha(rec['retained_backup']), sha(rec['destination']))
        self.assertFalse(self.stage.parent.stat().st_file_attributes & 2)

    def test_referenced_and_changed_copies_stay_in_place(self):
        save(self.plan.parent/'receipt.json', {'backup': str(self.source(2))})
        self.source(1).write_bytes(b'foreign content')
        result = coordinated.complete(self.journal, self.state)
        self.assertEqual(result['status'], 'PASS')
        self.assertEqual(core.journal_read(self.journal)['phase'], 'COMPLETE')
        self.assertEqual([r['path'] for r in result['service_archive']['retained']],
                         [str(self.source(1)), str(self.source(2))])
        self.assertEqual(self.source(1).read_bytes(), b'foreign content')
        self.assertTrue(self.source(2).is_file())

    def test_windows_busy_copy_does_not_block_completion_or_other_roles(self):
        import win32con
        import win32file
        busy = self.source(1)
        handle = win32file.CreateFile(str(busy), win32con.GENERIC_READ,
                                      win32con.FILE_SHARE_READ, None, win32con.OPEN_EXISTING, 0, None)
        try:
            result = coordinated.complete(self.journal, self.state)
        finally:
            handle.Close()
        self.assertEqual(result['status'], 'PASS')
        self.assertEqual(len(result['service_archive']['moved']), 2)
        self.assertEqual(result['service_archive']['retained'][0]['path'], str(busy))
        self.assertTrue(busy.is_file())

    def test_another_active_journal_reference_keeps_copy(self):
        busy = self.source(1)
        other = self.stage.parent / ('e'*32) / 'journal.json'
        other_plan = self.plan.parents[1] / ('e'*32) / 'plan.json'
        save(other_plan, {'inputs': {str(busy): sha(busy)}})
        core.journal_write(other, {'phase': 'STARTED', 'plan': str(other_plan)})
        before = sha(other)
        result = coordinated.complete(self.journal, self.state)['service_archive']
        self.assertEqual(len(result['moved']), 1)  # XLSX name is shared; fail closed.
        self.assertTrue(busy.is_file())
        self.assertEqual(sha(other), before)

    def test_interrupted_rename_audit_can_be_verified_without_second_move(self):
        item = self.state['participants'][1]
        source = self.source(1)
        destination = self.stage/'manual-archive'/source.name
        real = archive._record
        def interrupted(path, value):
            if path.name == 'move-verified.json':
                raise OSError('injected interruption after rename')
            real(path, value)
        with patch.object(archive, '_record', side_effect=interrupted):
            with self.assertRaisesRegex(OSError, 'injected'):
                archive.move_verified(source, destination, item['backup'], item['original_sha256'])
        self.assertFalse(source.exists())
        self.assertEqual(sha(destination), item['original_sha256'])
        archive.move_verified(source, destination, item['backup'], item['original_sha256'])
        self.assertEqual(read(destination.parent/'move-verified.json')['source'], str(source))

    def test_collision_does_not_replace_existing_archive(self):
        item = self.state['participants'][1]
        source = self.source(1)
        destination = self.stage/'already-archived.xlsx'
        destination.write_bytes(b'independent file')
        with self.assertRaises(Stop):
            archive.move_verified(source, destination, item['backup'], item['original_sha256'])
        self.assertEqual(destination.read_bytes(), b'independent file')
        self.assertEqual(sha(source), item['original_sha256'])

    def test_failed_post_move_sha_restores_original_path(self):
        item = self.state['participants'][1]
        source = self.source(1)
        destination = self.stage/'manual-archive'/source.name
        def failed_readback(path):
            return '0'*64 if Path(path) == destination else sha(path)
        with patch.object(archive, 'sha', side_effect=failed_readback):
            with self.assertRaisesRegex(Stop, 'SHA-256'):
                archive.move_verified(source, destination, item['backup'], item['original_sha256'])
        self.assertFalse(destination.exists())
        self.assertEqual(sha(source), item['original_sha256'])
        self.assertFalse((destination.parent/'move-verified.json').exists())


@unittest.skipUnless(sys.platform == 'win32', 'Windows replacement and locking')
class CatalogAndCrashTests(unittest.TestCase):
    def setUp(self):
        temp = TemporaryDirectory()
        self.addCleanup(temp.cleanup)
        self.root = Path(temp.name)

    def test_new_project_migration_and_refresh_keep_two_verified_backups(self):
        project = 'Новый проект'
        folder = self.root/project
        folder.mkdir()
        registry = self.root/'registry.xlsx'
        b = Workbook()
        headers = list(dict.fromkeys(h.casefold() for h in (*REQUIRED.values(), *model.HEADERS.values())))
        data = dict(zip(REQUIRED.values(), [project, '1/2026', 100, datetime(2026, 11, 11), 'C', project]))
        b.active.append(headers)
        b.active.append([data.get(h) for h in headers])
        b.active.add_table(Table(displayName='Registry', ref=f'A1:{get_column_letter(len(headers))}2'))
        b.save(registry)
        b.close()
        catalog = folder/'Марки.xlsx'
        b = Workbook()
        b.active.append(['Марка','Профиль','Краска','№ АКЗ','Примечание'])
        b.active.append(['M','16П','PAINT',0,'KEEP'])
        b.save(catalog)
        b.close()
        cfg = {'registry': str(registry), 'output': str(self.root),
               'first_passport_documents': str(self.root)}
        original = sha(catalog)
        assignments = catalog_ledger.read_assignments(catalog)
        migrated = catalog_ledger.migrate_in_place(catalog, cfg, project)
        migrated_sha = sha(catalog)
        refreshed = catalog_ledger.migrate_in_place(catalog, cfg, project)
        for res, expected in ((migrated, original), (refreshed, migrated_sha)):
            backup = Path(res['backup'])
            self.assertEqual(sha(backup), expected)
            self.assertIn('.vzmk-service-archive', backup.parts)
            self.assertEqual(len(res['service_archive']['moved']), 1, res)
        self.assertEqual(catalog_ledger.read_assignments(catalog), assignments)
        book = load_workbook(catalog)
        try:
            self.assertEqual(catalog_ledger.metadata(book)['project'], project)
            self.assertEqual(book[catalog_ledger.SOURCE]['E2'].value, 'KEEP')
            self.assertEqual(book.worksheets[0].max_column, 6)
        finally:
            book.close()
        self.assertEqual({p.name for p in folder.iterdir()}, {catalog.name, '.vzmk-service-archive'})

    def test_failed_migration_keeps_original_and_mandatory_backup(self):
        catalog = self.root/'Marks.xlsx'
        b = Workbook()
        b.active.append(['Марка','Профиль','Краска','№ АКЗ'])
        b.active.append(['M','16П','PAINT',0])
        b.save(catalog)
        b.close()
        original = sha(catalog)
        with patch.object(catalog_ledger, 'registered_history', return_value=[]):
            with patch('passport_app.excel.atomic_replace', side_effect=OSError('injected')):
                with self.assertRaisesRegex(OSError, 'injected'):
                    catalog_ledger.migrate_in_place(catalog, {}, 'Project')
        self.assertEqual(sha(catalog), original)
        backups = list((self.root/'.vzmk-service-archive').rglob('*.before-ledger-*.xlsx'))
        self.assertEqual(len(backups), 1)
        self.assertEqual(sha(backups[0]), original)

    def test_real_process_exit_during_commit_preserves_restore_inputs(self):
        project = self.root/'project'
        project.mkdir()
        run = self.root/'run'
        run.mkdir()
        id_ = 'a'*32
        inputs = {}
        participants = {}
        for role, folder, suffix in [('registry', self.root, '.xlsx'),
                                      ('catalog', project, '.xlsx'), ('uzk', self.root, '.xlsm')]:
            p, c = folder/(role+suffix), run/(role+suffix)
            p.write_bytes(('original '+role).encode())
            c.write_bytes(('updated '+role).encode())
            inputs[str(p)] = sha(p)
            participants[role] = {'role': role, 'path': str(p), 'candidate': str(c),
                                  'original_sha256': sha(p), 'candidate_sha256': sha(c)}
        registry = participants['registry']
        target = project/'2_2026'
        output = run/'2_2026.pdf'
        output.write_bytes(b'test document')
        request = [dict(id=id_, registry=registry['path'], target=str(target), run=str(run),
                        project='P', number='2/2026', fingerprints=inputs, cfg={'training': True}),
                   {'registry': registry['candidate'], 'snapshot_sha256': registry['candidate_sha256'],
                    'accounting': [participants['catalog'], participants['uzk']]},
                   {'outputs': [{'path': str(output)}], 'fingerprints': {str(output): sha(output)}},
                   {'created_at': 'test'}, {}]
        data, prepared, released, confirmation, _ = request
        data['seal'] = digest(data)
        save(run/'verified.json', {'plan_seal': data['seal']})
        prepared['plan_seal'] = released['plan_seal'] = confirmation['plan_seal'] = data['seal']
        confirmation['verified_sha256'] = sha(run/'verified.json')
        for name, value in (('plan', data), ('prepared', prepared), ('release', released),
                            ('confirmation', confirmation)):
            save(run/(name+'.json'), value)
        proof = run/'request.json'
        save(proof, request)
        child = '''
import os,sys
from production_app import coordinated as c
from production_app.storage import read
real = c.atomic_replace
def interrupted(*a, **k):
    real(*a, **k)
    if a[1].endswith('catalog.xlsx'):
        os._exit(92)
c.atomic_replace = interrupted
c.commit(*read(sys.argv[1]))
'''
        repo = Path(coordinated.__file__).parents[1]
        proc = subprocess.run([sys.executable, '-B', '-c', child, str(proof)], cwd=repo, timeout=30)
        self.assertEqual(proc.returncode, 92)
        journal = self.root/'.vzmk-operations'/id_/'journal.json'
        state = core.journal_read(journal)
        self.assertEqual(state['phase'], 'REPLACE_INTENT_CATALOG')
        self.assertTrue(archive.replacement_path(registry['path'], id_).exists())
        self.assertTrue(archive.replacement_path(participants['catalog']['path'], id_).exists())
        self.assertFalse((journal.parent/'service-archive').exists())
        self.assertEqual(core.recover(journal)['phase'], 'ROLLED_BACK')
        self.assertTrue(all(sha(p) == h for p, h in inputs.items()))
        self.assertFalse(target.exists())
        self.assertEqual(sha(journal.parent/'unpublished_documents'/output.name), sha(output))
