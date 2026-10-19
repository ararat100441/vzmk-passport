"""Isolated core.release integration: one print PDF before any registry change."""
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from pypdf import PdfWriter
from production_app import core
from production_app.storage import pins, read, save, sha, Stop


class ReleasePrintIntegration(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.root = Path(self.tmp.name)
        self.run = self.root / 'private-job'
        self.run.mkdir()
        self.name = '9_1_2026 от 19.10.2026'
        self.target = self.root / 'published' / self.name
        self.registry = self.root / 'registry.xlsx'
        self.registry.write_bytes(b'ORIGINAL_REGISTRY')
        self.registry_hash = sha(self.registry)
        prepared_root = self.root / 'prepared'
        prepared_root.mkdir()
        roles = ['appendix','quality','vik','uzk','akz0','akz_appendix0']
        self.sources = []
        for role in roles:
            suffix = '.docx' if role in ('quality','vik','uzk') else '.xlsx'
            file = prepared_root / (role + suffix)
            file.write_bytes(('SOURCE-' + role).encode('ascii'))
            self.sources.append({'role':role,'path':str(file)})
        self.sources_hash = pins([item['path'] for item in self.sources])
        save(self.run/'prepared.json', {'outputs':self.sources})
        save(self.run/'verified.json', {'fingerprints':self.sources_hash})
        self.plan={'run':str(self.run),'seal':'PLAN-SEAL','target':str(self.target),
                   'cfg':{'training':True},'number':'9/1/2026'}
        save(self.run/'confirmation.json',{'plan_seal':'PLAN-SEAL',
             'verified_sha256':sha(self.run/'verified.json')})
        self.plan_path = self.run / 'plan.json'
        self.plan_path.write_text('SYNTHETIC')

    def tearDown(self):
        self.tmp.cleanup()

    @staticmethod
    def fake_export(items):
        for item in items:
            pdf=PdfWriter()
            pdf.add_blank_page(width=595,height=842)
            with item['pdf'].open('wb') as out:pdf.write(out)
            pdf.close()

    def run_release(self,fail=False):
        with patch('production_app.core.load_plan',return_value=self.plan), \
             patch('production_app.office.verify_documents',return_value=[]), \
             patch('production_app.print_pdf._export_word',side_effect=self.fake_export), \
             patch('production_app.print_pdf._export_excel',
                   side_effect=Stop('PDF export failed') if fail else self.fake_export):
            return core.release(self.plan_path)

    def test_automatic_final_print_pdf_is_part_of_sealed_release(self):
        result = self.run_release()
        self.assertEqual(result['status'],'PASS')
        self.assertEqual(result['print_pdf'],self.name+'.pdf')
        info=read(self.run/'release.json')
        self.assertEqual([section['role'] for section in info['print_pdf']['sections']],
                         ['quality','appendix','vik','uzk','akz0','akz_appendix0'])
        files={item['role']:Path(item['path']) for item in info['outputs']}
        pdf=files['print_pdf']
        self.assertEqual(pdf.parent,self.run/'release')
        self.assertEqual(pdf.name,self.name+'.pdf')
        self.assertEqual(info['fingerprints'][str(pdf)],sha(pdf))
        self.assertFalse(self.target.exists())
        self.assertEqual(self.registry_hash,sha(self.registry))
        self.assertEqual(self.sources_hash,pins([item['path'] for item in self.sources]))

    def test_failed_printing_never_produces_releasable_manifest(self):
        with self.assertRaisesRegex(Stop,'PDF export failed'):
            self.run_release(fail=True)
        self.assertFalse((self.run/'release.json').exists())
        self.assertFalse((self.run/'release'/ (self.name+'.pdf')).exists())
        self.assertFalse(self.target.exists())
        self.assertEqual(self.registry_hash,sha(self.registry))
        self.assertEqual(self.sources_hash,pins([item['path'] for item in self.sources]))


if __name__ == '__main__':
    unittest.main(verbosity=2)
