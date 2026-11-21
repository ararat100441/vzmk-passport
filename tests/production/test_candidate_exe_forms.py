"""Optional Windows end-to-end check of an explicit frozen candidate, without working data."""
import hashlib
import json
import os
from pathlib import Path
import shutil
import subprocess
import tempfile
import unittest

from docx import Document
from openpyxl import load_workbook


@unittest.skipUnless(os.name == 'nt' and os.environ.get('VZMK_TEST_EXE'),
                     'Requires Windows, Word, Excel and VZMK_TEST_EXE')
class FrozenOfficeForms(unittest.TestCase):
    def test_create_and_reopen_forms_in_disposable_folder(self):
        exe = Path(os.environ['VZMK_TEST_EXE']).resolve(strict=True)
        with exe.open('rb') as stream:
            binary_hash = hashlib.file_digest(stream, 'sha256').hexdigest()
        root = Path(tempfile.mkdtemp(prefix='vzmk-exe-'))
        forms = root / 'forms'
        req, result = root / 'request.json', root / 'result.json'
        req.write_text(json.dumps({'folder': str(forms)}), encoding='utf8')
        env = dict(os.environ, LOCALAPPDATA=str(root/'profile'))
        # On timeout keep the evidence and allow the supervisor to clean up Office.
        proc = subprocess.Popen([str(exe), '--task', 'forms', '--request', str(req),
                                 '--result', str(result)], cwd=root, env=env)
        print(f'EXE SHA-256 {binary_hash}, test files {root}', flush=True)
        self.assertEqual(proc.wait(timeout=600), 0)
        data = json.loads(result.read_text(encoding='utf8'))
        self.assertEqual(data['status'], 'PASS')
        proof = json.loads(Path(data['office_supervisor']).read_text(encoding='utf8'))
        self.assertEqual((proof['owned_orphans'], proof['ownership_errors']), (0, []))
        self.assertGreaterEqual(len(proof['instances']), 2)
        for instance in proof['instances']:
            self.assertEqual(instance['exit_code'], 0)
            self.assertFalse(instance.get('forced_termination_attempted'))
        expected = {n+'.docx' for n in ('vik','uzk','quality')}
        expected.update(f'{role}{i}.xlsx' for role in ('akz','akz_appendix') for i in range(4))
        self.assertEqual({p.name for p in forms.iterdir()}, expected)
        for path in forms.iterdir():
            if path.suffix == '.docx':
                self.assertIn('{{NUMBER}}', '\n'.join(p.text for p in Document(path).paragraphs))
            else:
                book = load_workbook(path, read_only=True)
                try:
                    self.assertIn('{{NUMBER}}', book['АКЗ']['A1'].value)
                finally:
                    book.close()
        shutil.rmtree(root)
