"""Portable tests for transaction orchestration with Windows locks/COM replaced at their boundary."""
from contextlib import nullcontext
from copy import deepcopy
from pathlib import Path
from tempfile import TemporaryDirectory
import os
import shutil
import unittest
from unittest.mock import Mock, patch

from passport_app import excel
from production_app import coordinated, core, model, service_archive
from production_app.storage import Stop, sha


class RecoveryBoundaries(unittest.TestCase):
    def setUp(self):
        directory = TemporaryDirectory()
        self.addCleanup(directory.cleanup)
        self.root = Path(directory.name)
        self.run = self.root/'run'
        self.run.mkdir()
        self.token = '1'*32
        self.journal = self.root/'.vzmk-operations'/self.token/'journal.json'
        self.registry = self.root/'registry.xlsx'
        self.target = self.root/'documents'/'2_2026'

        self.replace(coordinated, 'mutex', lambda *args, **kwargs: nullcontext())
        self.replace(coordinated, 'locked_file', lambda path, **kwargs: nullcontext(Path(path).read_bytes()))
        self.replace(coordinated, 'frozen', lambda *args: nullcontext())
        self.replace(model, 'visible_vom', lambda path: object())
        self.replace(core, 'verify_registry', lambda *args, **kwargs: None)
        self.replace(coordinated, 'atomic_replace', self.replace_file)

    def replace(self, module, name, callback):
        p = patch.object(module, name, callback)
        p.start()
        self.addCleanup(p.stop)

    def replace_file(self, candidate, current, *, backup):
        self.assertEqual(Path(backup).parent.stat().st_dev, Path(current).parent.stat().st_dev)
        if Path(backup).exists():
            raise FileExistsError(backup)
        shutil.copyfile(current, backup)
        os.replace(candidate, current)

    def issue(self, *, colocated=False, catalog_parent=None):
        parent = self.root if colocated else (catalog_parent or self.root/'catalog')
        parent.mkdir(exist_ok=True)
        items = []
        for role, path in [('registry', self.registry), ('catalog', parent/'stock.xlsx'), ('uzk', self.root/'УЗК.xlsm')]:
            candidate = self.run/path.name
            path.write_bytes(b'old '+role.encode())
            candidate.write_bytes(b'new '+role.encode())
            items.append(dict(role=role, path=str(path), candidate=str(candidate),
                              original_sha256=sha(path), candidate_sha256=sha(candidate)))
        pdf = self.run/'2_2026.pdf'
        pdf.write_bytes(b'fixture')
        data = dict(id=self.token, registry=str(self.registry), target=str(self.target), run=str(self.run),
                    project='Project', number='2/2026', fingerprints={i['path']: i['original_sha256'] for i in items},
                    cfg={}, date='2026-11-20', vom='test')
        prepared = dict(registry=items[0]['candidate'], snapshot_sha256=items[0]['candidate_sha256'], accounting=items[1:])
        released = dict(outputs=[{'path': str(pdf)}], fingerprints={str(pdf): sha(pdf)})
        coordinated.commit(data, prepared, released, {'created_at':'test'}, {})
        return core.journal_read(self.journal)

    def test_registry_and_catalog_can_share_directory(self):
        state = self.issue(colocated=True)
        self.assertEqual(state['phase'], 'NEEDS_READBACK')
        paths = service_archive.replacement_paths(state['participants'], self.token)
        self.assertEqual(len(set(paths)), 3)
        for item, backup in zip(state['participants'], paths):
            self.assertEqual(sha(backup), item['original_sha256'])
            self.assertEqual(sha(item['path']), item['candidate_sha256'])
        self.assertEqual(core.recover(self.journal)['phase'], 'ROLLED_BACK')

    def test_stale_recovery_cannot_reverse_complete(self):
        stale = self.issue()
        current = deepcopy(stale)
        current['phase'] = 'COMPLETE'
        core.journal_write(self.journal, current)
        with self.assertRaises(Stop):
            coordinated.recover(self.journal, stale)
        self.assertEqual(core.journal_read(self.journal), current)
        self.assertTrue(self.target.is_dir())
        for item in stale['participants']:
            self.assertEqual(sha(item['path']), item['candidate_sha256'])

    def test_stale_completion_cannot_replace_newer_journal(self):
        stale = self.issue()
        newer = deepcopy(stale)
        newer['phase'] = 'COMPENSATING'
        core.journal_write(self.journal, newer)
        with self.assertRaises(Stop):
            coordinated.complete(self.journal, stale)
        self.assertEqual(core.journal_read(self.journal), newer)

    def test_unpublish_denied_preserves_all_accounting_candidates(self):
        state = self.issue()
        original = Path.rename

        def refuse(source, destination):
            if source == self.target:
                raise PermissionError('folder in use')
            return original(source, destination)

        with patch.object(Path, 'rename', refuse), self.assertRaises(PermissionError):
            core.recover(self.journal)
        self.assertTrue(self.target.is_dir())
        for item in state['participants']:
            self.assertEqual(sha(item['path']), item['candidate_sha256'])
        self.assertEqual(core.recover(self.journal)['phase'], 'ROLLED_BACK')

    def test_distinct_volumes_recover_documents_and_catalog(self):
        second = os.getenv('VZMK_SECOND_TEST_VOLUME')
        if not second:
            self.skipTest('Set VZMK_SECOND_TEST_VOLUME to a writable directory on another volume')
        directory = TemporaryDirectory(dir=second)
        self.addCleanup(directory.cleanup)
        other = Path(directory.name)
        self.assertNotEqual(other.stat().st_dev, self.root.stat().st_dev)
        self.target = other/'2_2026'
        state = self.issue(catalog_parent=other)
        self.assertEqual(core.recover(self.journal)['phase'], 'ROLLED_BACK')
        self.assertFalse(self.target.exists())
        archive = other/'.vzmk-service-archive'/self.token/'recovery'
        self.assertTrue((archive/'unpublished_documents'/'2_2026.pdf').is_file())
        self.assertEqual(len(list(archive.glob('pre-restore-*.xlsx'))), 1)
        for item in state['participants']:
            self.assertEqual(sha(item['path']), item['original_sha256'])

    def test_interrupted_after_first_restore_can_resume_from_sealed_backups(self):
        state = self.issue()
        calls = []
        def interrupt_once(candidate, path, *, backup):
            self.replace_file(candidate, path, backup=backup)
            calls.append(path)
            if len(calls) == 1:
                raise OSError('interruption after replacement')
        with patch.object(coordinated, 'atomic_replace', interrupt_once), self.assertRaises(OSError):
            core.recover(self.journal)
        self.assertEqual(core.recover(self.journal)['phase'], 'ROLLED_BACK')
        self.assertEqual(core.recover(self.journal)['phase'], 'ROLLED_BACK')
        self.assertEqual(core.pending(self.registry), [])
        archived = self.journal.parent/'unpublished_documents'
        self.assertEqual(sha(archived/'2_2026.pdf'), state['output_hashes']['2_2026.pdf'])
        for item in state['participants']:
            self.assertEqual(sha(item['path']), item['original_sha256'])

    def test_interruption_after_unpublish_rechecks_archived_documents(self):
        state = self.issue()
        with patch.object(coordinated, '_write_bytes', side_effect=OSError('abrupt interruption')), self.assertRaises(OSError):
            core.recover(self.journal)
        self.assertFalse(self.target.exists())
        archived = self.journal.parent/'unpublished_documents'/'2_2026.pdf'
        original = archived.read_bytes()
        archived.write_bytes(b'external document correction')
        with self.assertRaises(Stop):
            core.recover(self.journal)
        for item in state['participants']:
            self.assertEqual(sha(item['path']), item['candidate_sha256'])
        archived.write_bytes(original)
        self.assertEqual(core.recover(self.journal)['phase'], 'ROLLED_BACK')

    def test_missing_published_folder_and_archive_keep_journal_pending(self):
        state = self.issue()
        self.target.rename(self.target.with_name('moved outside the application'))
        with self.assertRaises(Stop):
            core.recover(self.journal)
        self.assertEqual(core.journal_read(self.journal), state)
        for item in state['participants']:
            self.assertEqual(sha(item['path']), item['candidate_sha256'])

    def test_newer_external_bytes_prevent_document_and_accounting_rollback(self):
        state = self.issue()
        item = state['participants'][1]
        Path(item['path']).write_bytes(b'newer external stock')
        with self.assertRaises(Stop):
            core.recover(self.journal)
        self.assertEqual(Path(item['path']).read_bytes(), b'newer external stock')
        self.assertTrue(self.target.is_dir())
        self.assertEqual(sha(self.registry), state['participants'][0]['candidate_sha256'])
        self.assertEqual(core.journal_read(self.journal), state)

    @unittest.skipUnless(os.name == 'nt', 'Native Windows locking and ReplaceFileW')
    def test_native_colocated_commit_rollback(self):
        from production_app.storage import mutex
        from stage8_closure.storage import frozen
        with patch.multiple(coordinated, locked_file=excel.locked_file, mutex=mutex,
                            atomic_replace=excel.atomic_replace, frozen=frozen):
            state = self.issue(colocated=True)
            self.assertEqual(core.recover(self.journal)['phase'], 'ROLLED_BACK')
        self.assertTrue((self.journal.parent/'unpublished_documents'/'2_2026.pdf').is_file())
        for item in state['participants']:
            self.assertEqual(sha(item['path']), item['original_sha256'])

    def test_windows_local_unc_and_extended_names(self):
        samples = [(r'C:\folder\f.xlsx', r'\\?\C:\folder\f.xlsx'),
                   (r'\\server\share\f.xlsx', r'\\?\UNC\server\share\f.xlsx'),
                   (r'\\?\UNC\server\share\f.xlsx', r'\\?\UNC\server\share\f.xlsx')]
        for source, expected in samples:
            with self.subTest(source=source):
                kernel = Mock()
                kernel.ReplaceFileW.return_value = 1
                with patch.object(excel.ctypes, 'WinDLL', return_value=kernel, create=True), \
                        patch.object(Path, 'resolve', lambda path: path):
                    excel.atomic_replace(source, source, backup=source)
                self.assertEqual(kernel.ReplaceFileW.call_args.args[:3], (expected,) * 3)
