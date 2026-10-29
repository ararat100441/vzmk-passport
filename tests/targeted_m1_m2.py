"""Narrow M1/M2 checks; all writes confined to the supplied validation directory."""
import argparse
import importlib.util
import json
import shutil
import subprocess
import sys
import unittest
from pathlib import Path
from unittest.mock import patch

BASE = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(BASE))
from passport_app import vom
from passport_app.excel import load_book, sha256
from passport_app.tables import locate_layout

spec = importlib.util.spec_from_file_location('finalize_evidence', BASE / 'tests/finalize_evidence.py')
finalizer = importlib.util.module_from_spec(spec)
spec.loader.exec_module(finalizer)
AUDIT = json.loads((BASE / 'audit/audit_inventory.json').read_text(encoding='utf-8'))
SOURCE_VOM = Path(AUDIT['vom']['path'])


class ConnectionsTests(unittest.TestCase):
    def parse_names(self, names):
        formulas, values = load_book(SOURCE_VOM), load_book(SOURCE_VOM, values=True)
        layout = locate_layout(formulas)
        sheet = values[layout.sheet]
        # Only the in-memory cell names differ; no source workbook is saved.
        for r in range(layout.start, sheet.max_row + 1):
            sheet.cell(r, layout.columns['name']).value = None
        for r, name in enumerate(names, layout.start):
            sheet.cell(r, layout.columns['name']).value = name
        with patch.object(vom, 'load_book', side_effect=[formulas, values]):
            return vom.read_vom(SOURCE_VOM)

    def test_normal(self):
        self.assertEqual(self.parse_names(['Балка', 'Балка']).connections, 'Балка')

    def test_edge_spaces_preserved(self):
        self.assertEqual(self.parse_names([' Балка ']).connections, ' Балка ')

    def test_spaces_define_unique_values_in_first_order(self):
        result = self.parse_names([' Балка ', 'Балка', ' Балка '])
        self.assertEqual(result.connections, ' Балка , Балка')
        self.assertEqual([r[1] for r in result.rows[:3]], [' Балка ', 'Балка', ' Балка '])

    def test_empty_and_whitespace_skipped(self):
        self.assertEqual(self.parse_names([None, '', '   ', '\t ', 'Балка']).connections, 'Балка')

    def test_real_control(self):
        result = vom.read_vom(SOURCE_VOM)
        self.assertEqual(result.connections, 'Балка')
        self.assertEqual((len(result.rows), str(result.quantity), str(result.mass), str(result.area)),
                         (32, '50', '15071.44', '477.749'))


def copy_outputs(destination):
    destination.mkdir()
    for name in finalizer.CASES:
        original = finalizer.ROOT / name
        copy = destination / name
        shutil.copytree(original, copy)
        for journal in (copy / '.pasport-transactions').glob('*/journal.json'):
            data = json.loads(journal.read_text(encoding='utf-8'))
            for key in ('registry', 'backup', 'passport', 'pending_folder'):
                data[key] = str(copy / Path(data[key]).relative_to(original))
            journal.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding='utf-8')
    return destination


class EvidenceTests(unittest.TestCase):
    def test_pass_requires_three_actual_verifications(self):
        root = copy_outputs(OUT / 'm2_pass')
        with patch.object(finalizer, 'verify_saved', wraps=finalizer.verify_saved) as verifier:
            report = finalizer.verify_outputs(root)
        self.assertEqual(verifier.call_count, 3)
        self.assertTrue(report['pass'])
        self.assertTrue(all(c['pass'] and c['verify_saved_called'] for c in report['cases']))
        for item, call in zip(report['cases'], verifier.call_args_list):
            for role, argument in [('registry', 0), ('backup', 1), ('appendix', 2)]:
                actual = Path(call.args[argument])
                self.assertEqual(item['files'][role], {'path': str(actual), 'sha256': sha256(actual)})

    def test_failed_verifier_overwrites_stale_pass_and_exits_nonzero(self):
        root = copy_outputs(OUT / 'm2_fail')
        evidence = root / 'saved_guard_results.json'
        evidence.write_text('{"pass": true}', encoding='utf-8')
        command = (
            "import sys; sys.path.insert(0, sys.argv[1]); import finalize_evidence as f; "
            "f.verify_saved=lambda *a, **k: (_ for _ in ()).throw(RuntimeError('injected verify_saved failure')); "
            "f.main(['--verify-only', '--root', sys.argv[2]])"
        )
        process = subprocess.run([sys.executable, '-c', command, str(BASE / 'tests'), str(root)],
                                 capture_output=True, text=True, encoding='utf-8', errors='replace')
        report = json.loads(evidence.read_text(encoding='utf-8'))
        self.assertNotEqual(process.returncode, 0)
        self.assertFalse(report['pass'])
        self.assertEqual(len(report['cases']), 3)
        self.assertTrue(all(not c['pass'] and c['verify_saved_called'] and
                            'injected verify_saved failure' in c['error'] for c in report['cases']))
        (OUT / 'm2_fail_process.json').write_text(json.dumps({
            'returncode': process.returncode, 'stderr': process.stderr,
            'saved_guard_results': str(evidence), 'false_pass_created': report['pass'],
        }, ensure_ascii=False, indent=2), encoding='utf-8')


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--out', type=Path, required=True)
    OUT = parser.parse_args().out.resolve()
    if not OUT.is_relative_to((BASE / 'validation').resolve()) or not OUT.is_dir():
        raise ValueError('Use an existing validation directory')
    before = json.loads((OUT / 'before.json').read_text(encoding='utf-8'))
    sources_before = {p: sha256(p) for p in before['production_before']}
    if sources_before != before['production_before']:
        raise ValueError('Production differs from accepted checkpoint')
    suite = unittest.TestSuite([
        unittest.defaultTestLoader.loadTestsFromTestCase(ConnectionsTests),
        unittest.defaultTestLoader.loadTestsFromTestCase(EvidenceTests),
    ])
    result = unittest.TextTestRunner(verbosity=2).run(suite)
    sources_after = {p: sha256(p) for p in sources_before}
    report = {'pass': result.wasSuccessful() and sources_before == sources_after,
              'tests_run': result.testsRun,
              'failures': [str(test) + ': ' + message for test, message in result.failures + result.errors],
              'real_control': {'connections': 'Балка', 'rows': 32, 'quantity': 50,
                               'mass_kg': '15071.44', 'area_m2': '477.749'},
              'source_sha256_before': sources_before, 'source_sha256_after': sources_after,
              'm2_pass': str(OUT / 'm2_pass/saved_guard_results.json'),
              'm2_fail': str(OUT / 'm2_fail/saved_guard_results.json'),
              'm2_fail_process': str(OUT / 'm2_fail_process.json'),
              'full_regression': 'NOT_RUN'}
    (OUT / 'targeted_results.json').write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding='utf-8')
    sys.exit(0 if report['pass'] else 1)
