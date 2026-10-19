"""Synthetic merge/order/safety checks; native Microsoft Office not required."""
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from pypdf import PdfReader, PdfWriter
from production_app.print_pdf import (
    ordered_sources, create_print_pdf, verify_print_pdf, verify_package_print
)
from production_app.storage import Stop, sha

PAGE_COUNTS = {
    'quality': 2, 'appendix': 3, 'vik': 4, 'uzk': 5,
    'akz0': 6, 'akz_appendix0': 7,
    'akz2': 8, 'akz_appendix2': 9,
}
PDF_ORDER = ('quality','appendix','vik','uzk','akz0','akz_appendix0')


class PrintPDF(unittest.TestCase):
    def setUp(self):
        self.folder = tempfile.TemporaryDirectory()
        self.release = Path(self.folder.name) / 'release'
        self.release.mkdir()
        self.name = '9_1_2026 от 19.10.2026'

    def tearDown(self):
        self.folder.cleanup()

    def outputs(self, roles=PDF_ORDER):
        result = []
        for role in roles:
            suffix = '.docx' if role in ('quality','vik','uzk') else '.xlsx'
            path = self.release / (role + suffix)
            path.write_bytes(('SYNTHETIC-' + role).encode('ascii'))
            result.append({'role':role,'path':str(path)})
        return result

    @staticmethod
    def export_synthetic(items):
        for item in items:
            writer = PdfWriter()
            # A unique page width identifies each source and proves actual
            # order after merge; no string/metadata assumptions necessary.
            width = 200 + list(PAGE_COUNTS).index(item['role']) * 10
            for _ in range(PAGE_COUNTS[item['role']]):
                writer.add_blank_page(width=width, height=300)
            with item['pdf'].open('wb') as file:
                writer.write(file)
            writer.close()

    def create(self, outputs, name=None):
        with patch('production_app.print_pdf._export_word', side_effect=self.export_synthetic), \
                patch('production_app.print_pdf._export_excel', side_effect=self.export_synthetic):
            return create_print_pdf(outputs,self.release,name or self.name)

    def test_exact_priority_order_even_with_shuffled_inputs(self):
        docs=self.outputs(tuple(reversed(PDF_ORDER)))
        self.assertEqual([item['role'] for item in ordered_sources(docs)],list(PDF_ORDER))

    def test_merge_real_pdf_files_and_named_output_in_same_folder(self):
        docs=self.outputs(tuple(reversed(PDF_ORDER)))
        before={d['path']:sha(d['path']) for d in docs}
        created=self.create(docs)
        dest=Path(created['output']['path'])
        self.assertEqual(dest.name,self.name+'.pdf')
        self.assertEqual(dest.parent,self.release)
        self.assertEqual(created['manifest']['pages'],sum(PAGE_COUNTS[r] for r in PDF_ORDER))
        self.assertEqual([r['role'] for r in created['manifest']['sections']],list(PDF_ORDER))
        self.assertEqual([r['start_page'] for r in created['manifest']['sections']],[1,3,6,10,15,21])
        self.assertEqual(before,{d['path']:sha(d['path']) for d in docs})
        self.assertEqual({p.name for p in self.release.iterdir()},
                         {Path(x['path']).name for x in docs}|{dest.name})
        pdf=PdfReader(str(dest))
        widths=[int(p.mediabox.width) for p in pdf.pages]
        cursor=0
        for role in PDF_ORDER:
            group=widths[cursor:cursor+PAGE_COUNTS[role]]
            self.assertEqual(len(group),PAGE_COUNTS[role])
            self.assertEqual(set(group),{200+list(PAGE_COUNTS).index(role)*10})
            cursor+=len(group)
        release={'outputs':docs+[created['output']],'print_pdf':created['manifest']}
        self.assertEqual(verify_package_print(release,self.name),created['manifest'])

    def test_multiple_akz_variants_use_separate_act_and_appendix_pairs(self):
        roles=(*PDF_ORDER,'akz2','akz_appendix2')
        docs=self.outputs(roles)
        self.assertEqual([i['role'] for i in ordered_sources(docs)],list(roles))
        pdf=self.create(docs)
        self.assertEqual([s['role'] for s in pdf['manifest']['sections']],list(roles))
        self.assertEqual(pdf['manifest']['pages'],sum(PAGE_COUNTS[r] for r in roles))

    def test_no_akz_still_prints_available_approved_documents(self):
        docs=self.outputs(PDF_ORDER[:4])
        pdf=self.create(docs)
        self.assertEqual([s['role'] for s in pdf['manifest']['sections']],list(PDF_ORDER[:4]))

    def test_orphan_akz_must_not_silently_be_omitted(self):
        docs=self.outputs(PDF_ORDER)
        docs=[d for d in docs if d['role']!='akz_appendix0']
        with self.assertRaisesRegex(Stop,'акт и его приложение'):
            self.create(docs)
        self.assertFalse((self.release/(self.name+'.pdf')).exists())

    def test_missing_mandatory_document_fails_before_export(self):
        docs=[d for d in self.outputs(PDF_ORDER) if d['role']!='quality']
        with self.assertRaisesRegex(Stop,'отсутствуют'):
            self.create(docs)

    def test_duplicate_role_refused(self):
        docs=self.outputs()
        with self.assertRaisesRegex(Stop,'Повтор роли'):
            ordered_sources(docs+[docs[0]])

    def test_output_pdf_never_overwritten(self):
        docs=self.outputs()
        dest=self.release/(self.name+'.pdf')
        dest.write_bytes(b'original')
        with self.assertRaisesRegex(Stop,'перезапись'):
            self.create(docs)
        self.assertEqual(dest.read_bytes(),b'original')

    def test_tampering_pdf_fails_hash_verification(self):
        docs=self.outputs()
        result=self.create(docs)
        released={'outputs':docs+[result['output']],'print_pdf':result['manifest']}
        target=Path(result['output']['path'])
        target.write_bytes(target.read_bytes()+b'TAMPER')
        with self.assertRaisesRegex(Stop,'изменился'):
            verify_package_print(released,self.name)

    def test_incorrect_order_manifest_fails(self):
        docs=self.outputs()
        result=self.create(docs)
        released={'outputs':docs+[result['output']],'print_pdf':result['manifest']}
        released['print_pdf']['sections'][0]['role']='vik'
        with self.assertRaisesRegex(Stop,'порядок'):
            verify_package_print(released,self.name)

    def test_office_failure_creates_no_published_pdf(self):
        docs=self.outputs()
        before={d['path']:sha(d['path']) for d in docs}
        with patch('production_app.print_pdf._export_word',side_effect=RuntimeError('Office failure')):
            with self.assertRaisesRegex(RuntimeError,'Office failure'):
                create_print_pdf(docs,self.release,self.name)
        self.assertFalse((self.release/(self.name+'.pdf')).exists())
        self.assertEqual(before,{d['path']:sha(d['path']) for d in docs})

    def test_source_file_mutation_fails_before_merge(self):
        docs=self.outputs()
        def mutate(items):
            if items:
                items[0]['source'].write_bytes(b'altered')
            self.export_synthetic(items)
        with patch('production_app.print_pdf._export_word',side_effect=mutate), \
                patch('production_app.print_pdf._export_excel',side_effect=self.export_synthetic):
            with self.assertRaisesRegex(Stop,'SHA-256 изменился'):
                create_print_pdf(docs,self.release,self.name)
        self.assertFalse((self.release/(self.name+'.pdf')).exists())

    def test_pdf_pagecount_corruption_detected(self):
        docs=self.outputs()
        result=self.create(docs)
        with self.assertRaisesRegex(Stop,'количество страниц'):
            verify_print_pdf(result['output']['path'],expected_pages=123)

    def test_sources_outside_private_folder_refused(self):
        docs=self.outputs()
        alien=Path(self.folder.name)/'other.docx'
        alien.write_bytes(b'other')
        docs[0]['path']=str(alien)
        with self.assertRaisesRegex(Stop,'в одной папке'):
            self.create(docs)


if __name__ == '__main__':
    unittest.main(verbosity=2)
