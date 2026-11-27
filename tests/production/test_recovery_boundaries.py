"""Portable tests for transaction orchestration with Windows locks/COM replaced at their boundary."""
from contextlib import nullcontext
from copy import deepcopy
from pathlib import Path
from types import SimpleNamespace
from tempfile import TemporaryDirectory
import os
import shutil
import subprocess
import sys
import unittest
from unittest.mock import MagicMock, Mock, patch

from openpyxl import Workbook
from passport_app import excel
from pypdf import PdfWriter
from production_app import amendment, coordinated, core, model, service_archive
from production_app.storage import Stop, digest, save, sha


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

    def issue(self, *, colocated=False, catalog_parent=None, with_uzk_proof=False):
        parent = self.root if colocated else (catalog_parent or self.root/'catalog')
        parent.mkdir(exist_ok=True)
        items = []
        for role, path in [('registry', self.registry), ('catalog', parent/'stock.xlsx'), ('uzk', self.root/'УЗК.xlsm')]:
            candidate = self.run/path.name
            path.write_bytes(b'old '+role.encode())
            if role == 'uzk' and with_uzk_proof:
                source = Workbook()
                source.active.title = 'УЗК'
                source.active.append(['Заказ', 'Марка элемента', 'Кол-во (шт)'])
                source.active.append(['Project', 'M-1', 3])
                source.save(path)
                source.close()
            candidate.write_bytes(b'new '+role.encode())
            items.append(dict(role=role, path=str(path), candidate=str(candidate),
                              original_sha256=sha(path), candidate_sha256=sha(candidate)))
        pdf = self.run/'2_2026.pdf'
        documents = [('quality', 'quality.docx'), ('appendix', 'appendix.xlsx'),
                     ('vik', 'vik.docx'), ('uzk', 'uzk.docx')]
        writer = PdfWriter()
        for role, name in documents:
            (self.run/name).write_bytes(role.encode())
            writer.add_blank_page(width=595, height=842)
        writer.write(pdf)
        writer.close()
        data = dict(schema='VZMK_PRODUCTION_PLAN_1', id=self.token, registry=str(self.registry), target=str(self.target), run=str(self.run),
                    project='Project', number='2/2026', fingerprints={i['path']: i['original_sha256'] for i in items},
                    cfg={'registry': str(self.registry)}, date='2026-11-20', vom='test', facts={},
                    accounting={item['role']: item['path'] for item in items[1:]})
        if with_uzk_proof:
            data['accounting']['controls'] = {'M-1': 2}
        data['seal'] = digest(data)
        prepared = dict(plan_seal=data['seal'], registry=items[0]['candidate'],
                        snapshot_sha256=items[0]['candidate_sha256'], accounting=items[1:])
        outputs = [{'role': role, 'path': str(self.run/name)} for role, name in documents]
        outputs.append({'role': 'print_pdf', 'path': str(pdf)})
        fingerprints = {item['path']: sha(item['path']) for item in outputs}
        manifest = dict(filename=pdf.name, pages=len(documents), sha256=sha(pdf), sections=[
            {'role': role, 'pages': 1, 'start_page': index} for index, (role, _) in enumerate(documents, start=1)
        ])
        released = dict(plan_seal=data['seal'], outputs=outputs, fingerprints=fingerprints, print_pdf=manifest)
        for name, value in (('plan', data), ('prepared', prepared), ('release', released),
                            ('verified', {'plan_seal': data['seal']})):
            save(self.run/(name+'.json'), value)
        confirmation = {'created_at': 'test', 'reviewed': True, 'plan_seal': data['seal'],
                        'verified_sha256': sha(self.run/'verified.json')}
        save(self.run/'confirmation.json', confirmation)
        coordinated.commit(data, prepared, released, confirmation, {})
        return core.journal_read(self.journal)

    def test_archive_remains_authoritative_after_local_run_is_removed(self):
        state = self.issue()
        shutil.rmtree(self.run)
        evidence = service_archive.operation_evidence(self.journal, state)
        self.assertEqual(set(evidence), {'plan.json', 'prepared.json', 'verified.json',
                                         'confirmation.json', 'release.json'})
        self.assertTrue(all(path.parent == self.journal.parent/'evidence' for path in evidence.values()))
        self.assertEqual(core.recover(self.journal)['phase'], 'ROLLED_BACK')

    def test_changed_archive_is_rejected_even_with_local_original(self):
        state = self.issue()
        archived = self.journal.parent/'evidence'/'plan.json'
        save(archived, {'plan': 'changed'})
        self.assertTrue((self.run/'plan.json').is_file())
        with self.assertRaisesRegex(Stop, 'plan.json'):
            service_archive.operation_evidence(self.journal, state, names=('plan.json',))

    def test_readback_and_complete_use_archive_and_published_pdf_after_appdata_loss(self):
        state = self.issue(with_uzk_proof=True)
        shutil.rmtree(self.run)
        app = MagicMock()
        app.Workbooks.Open.return_value.Worksheets.return_value.Cells.return_value.Value2 = '2/2026'
        registry = SimpleNamespace(projects={'Project': [{'number': '2/2026', 'row': 2}]},
                                   sheet=1, columns={'number': 1})
        with (patch.object(core, 'excel_session') as excel_session,
              patch.object(core, 'read_registry', return_value=registry),
              patch.object(core.office, 'verify_documents'),
              patch.object(coordinated, 'native_readback') as native,
              patch.object(core, 'supervisor_proof') as supervisor):
            excel_session.return_value.__enter__.return_value = app
            result = core.readback(self.journal)
            self.assertEqual(result['plan'], state['plan'])
            self.assertEqual(result['print_pdf_pages'], 4)
            native.assert_called_once()
            result = core.complete(self.journal, self.root/'readback-result.json')
            supervisor.assert_called_once_with(self.root/'readback-result.json', 'readback', state['plan'])
        self.assertEqual(result['state'], 'ГОТОВО')
        self.assertEqual(core.pending(self.registry), [])
        self.assertEqual(len(result['service_archive']['moved']), 3 if os.name == 'nt' else 0)
        self.assertTrue(self.target.is_dir())
        uzk = self.root/'УЗК.xlsm'
        quantities, pins = amendment.refund_sources(self.registry, self.target, uzk, 'Project',
                                                     state['issue_sha256'], {'M-1': 1})
        self.assertEqual(quantities, {'M-1': {'project': 'Project', 'mark': 'M-1'}})
        self.assertEqual(set(map(Path, pins)), {self.journal, self.journal.parent/'evidence'/'plan.json',
                                               self.journal.parent/'backup-uzk.xlsm'})

    def test_low_destination_disk_space_stops_before_public_write(self):
        with patch.object(shutil, 'disk_usage', return_value=Mock(free=0)):
            with self.assertRaisesRegex(Stop, 'Недостаточно места'):
                self.issue()
        self.assertFalse(self.journal.exists())
        self.assertFalse(self.target.exists())
        self.assertEqual(self.registry.read_bytes(), b'old registry')

    def test_candidate_change_after_replace_keeps_intent_and_stops_next_replace(self):
        def replace_then_change(candidate, path, *, backup):
            self.replace_file(candidate, path, backup=backup)
            Path(path).write_bytes(b'external')

        with patch.object(coordinated, 'atomic_replace', side_effect=replace_then_change):
            with self.assertRaisesRegex(Stop, 'SHA'):
                self.issue()
        self.assertEqual(core.journal_read(self.journal)['phase'], 'REPLACE_INTENT_REGISTRY')
        self.assertEqual((self.root/'catalog'/'stock.xlsx').read_bytes(), b'old catalog')
        self.assertEqual((self.root/'УЗК.xlsm').read_bytes(), b'old uzk')

    @unittest.skipUnless(os.name == 'nt', 'Uses Windows rename without overwrite')
    def test_partial_replacement_can_restore_missing_name_only_with_verified_intent(self):
        from production_app.storage import mutex
        from stage8_closure.storage import frozen

        def partial(candidate, path, *, backup):
            Path(path).rename(backup)
            raise OSError('injected interrupted replace before installing updated workbook')

        with patch.multiple(coordinated, atomic_replace=partial, locked_file=excel.locked_file,
                            mutex=mutex, frozen=frozen):
            with self.assertRaisesRegex(Stop, 'injected interrupted replace'):
                self.issue()
        state = core.journal_read(self.journal)
        self.assertEqual(state['phase'], 'REPLACE_INTENT_REGISTRY')
        self.assertFalse(self.registry.exists())
        with patch.multiple(coordinated, locked_file=excel.locked_file, mutex=mutex,
                            atomic_replace=excel.atomic_replace, frozen=frozen):
            self.assertEqual(core.recover(self.journal)['phase'], 'ROLLED_BACK')
        self.assertEqual(self.registry.read_bytes(), b'old registry')
        replacement = service_archive.replacement_path(self.registry, self.token)
        self.assertEqual(replacement.read_bytes(), b'old registry')

    def test_unexplained_missing_live_workbook_does_not_restore_any_participant(self):
        state = self.issue()
        missing = Path(state['participants'][1]['path'])
        missing.rename(missing.with_name('external.xlsx'))
        with self.assertRaisesRegex(Stop, 'отсутствует'):
            core.recover(self.journal)
        self.assertTrue(self.target.is_dir())
        self.assertFalse(missing.exists())
        self.assertEqual(sha(self.registry), state['participants'][0]['candidate_sha256'])

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

    @unittest.skipUnless(os.name == 'nt', 'Native Windows locking and ReplaceFileW')
    def test_second_process_cannot_open_replaced_registry_during_next_participant_write(self):
        from production_app.storage import mutex
        from stage8_closure.storage import frozen

        worker = """
import sys
from passport_app.excel import locked_file
from passport_app.config import PassportError
from production_app.storage import mutex, Stop
try:
    with mutex(sys.argv[1]):
        raise AssertionError('registry mutex unexpectedly available')
except Stop:
    pass
try:
    with locked_file(sys.argv[1], writable=True):
        raise AssertionError('replaced registry file unexpectedly available')
except PassportError:
    pass
"""

        def inspect_before_catalog(candidate, path, *, backup):
            if Path(path).name == 'stock.xlsx':
                result = subprocess.run([sys.executable, '-B', '-c', worker, str(self.registry)],
                                        cwd=Path(coordinated.__file__).parents[1],
                                        capture_output=True, timeout=30)
                self.assertEqual(result.returncode, 0, result.stderr)
            excel.atomic_replace(candidate, path, backup=backup)

        with patch.multiple(coordinated, locked_file=excel.locked_file, mutex=mutex,
                            atomic_replace=inspect_before_catalog, frozen=frozen):
            self.assertEqual(self.issue()['phase'], 'NEEDS_READBACK')

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
