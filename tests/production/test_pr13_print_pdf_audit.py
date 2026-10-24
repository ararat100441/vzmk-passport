"""PR14 PDF diagnostic safeguards: 3+ pages must not silently PASS."""
import unittest
from types import SimpleNamespace
from unittest.mock import patch

from verify_appendix_print_com import (
    sample_rows, normalize_pdf_text, audit_pdf_pages,
)
from production_app.storage import Stop


class Page:
    def __init__(self,text,width=595,height=842):
        self._text=text
        self.mediabox=SimpleNamespace(width=width,height=height)
    def extract_text(self):
        return self._text


def pdf_pages(*texts):
    return SimpleNamespace(pages=[Page(t) for t in texts])


GOOD='Отправочная марка  Наименование марки   Количество  SYNTHETIC-001'


class MultiPagePdfAudit(unittest.TestCase):
    def test_diagnostic_includes_long_names_without_entering_working_registry(self):
        rows=sample_rows(110)
        self.assertEqual(len(rows),110)
        self.assertEqual(rows[0][0],'SYNTHETIC-001')
        self.assertEqual(rows[-1][0],'SYNTHETIC-110')
        self.assertGreater(len(rows[10][1]),len(rows[0][1])*2)
        self.assertEqual(rows[10][2],1)
        self.assertTrue(all(len(row)==7 for row in rows))

    def test_normalization_survives_wrap_between_header_words(self):
        self.assertIn('отправочнаямарка',
                      normalize_pdf_text('Отправочная \n марка'))

    def test_three_pages_with_full_headers_are_structurally_valid(self):
        with patch('verify_appendix_print_com.PdfReader',
                   return_value=pdf_pages(GOOD,GOOD,GOOD)):
            audit=audit_pdf_pages('unused.pdf','SYNTHETIC-001','SYNTHETIC-110')
        self.assertEqual(len(audit),3)
        self.assertTrue(all(p['has_mark_header'] for p in audit))
        self.assertTrue(all(p['has_name_header'] for p in audit))

    def test_missing_repeated_header_only_on_third_page_is_detected(self):
        with patch('verify_appendix_print_com.PdfReader',
                   return_value=pdf_pages(GOOD,GOOD,'SYNTHETIC-110')):
            with self.assertRaisesRegex(RuntimeError,'странице 3'):
                audit_pdf_pages('unused.pdf','SYNTHETIC-001','SYNTHETIC-110')

    def test_repeated_header_without_any_mark_rows_is_rejected(self):
        only_header='Отправочная марка  Наименование марки   Количество'
        with patch('verify_appendix_print_com.PdfReader',
                   return_value=pdf_pages(GOOD,GOOD,only_header)):
            with self.assertRaisesRegex(RuntimeError,'ни одной синтетической'):
                audit_pdf_pages('unused.pdf','SYNTHETIC-001','SYNTHETIC-110')

    def test_missing_page_text_is_detected(self):
        with patch('verify_appendix_print_com.PdfReader',
                   return_value=pdf_pages(GOOD,GOOD,'')):
            with self.assertRaisesRegex(RuntimeError,'не содержит извлекаемого'):
                audit_pdf_pages('unused.pdf','SYNTHETIC-001','SYNTHETIC-110')

    def test_too_few_pages_rejected_to_exercise_third_page(self):
        with patch('verify_appendix_print_com.PdfReader',
                   return_value=pdf_pages(GOOD,GOOD)):
            with self.assertRaisesRegex(RuntimeError,'третью|3-ю|3-'):
                audit_pdf_pages('unused.pdf','SYNTHETIC-001','SYNTHETIC-110')

    def test_broken_pdf_media_box_rejected(self):
        bad=pdf_pages(GOOD,GOOD,GOOD)
        bad.pages[2].mediabox.width=10
        with patch('verify_appendix_print_com.PdfReader',return_value=bad):
            with self.assertRaisesRegex(RuntimeError,'размер PDF страницы 3'):
                audit_pdf_pages('unused.pdf','SYNTHETIC-001','SYNTHETIC-110')


if __name__=='__main__':
    unittest.main(verbosity=2)
