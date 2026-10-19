"""Synthetic, read-only Word content-control contracts; no Office required."""
import tempfile
import unittest
import zipfile
from pathlib import Path

from production_app.storage import Stop
from production_app.word_controls import inspect_word_fields, fill_word_controls

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'


def tagged(tag, body='', placeholder=False):
    attr = '<w:showingPlcHdr/>' if placeholder else ''
    return (f'<w:sdt><w:sdtPr><w:tag w:val="{tag}"/>{attr}</w:sdtPr>'
            f'<w:sdtContent><w:r><w:t>{body}</w:t></w:r></w:sdtContent></w:sdt>')


class FakeControl:
    def __init__(self, tag, text='', showing=False, locked=False):
        self.Tag = tag
        self.Range = type('Range', (), {'Text': text})()
        self.ShowingPlaceholderText = showing
        self.LockContents = locked


class FakeCollection:
    def __init__(self, *controls):
        self.controls = controls
        self.Count = len(controls)

    def Item(self, index):
        return self.controls[index - 1]


class WordControlContracts(unittest.TestCase):
    def setUp(self):
        self.folder = tempfile.TemporaryDirectory()

    def tearDown(self):
        self.folder.cleanup()

    def doc(self, content):
        path = Path(self.folder.name) / 'template.docx'
        xml = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
               f'<w:document xmlns:w="{W}"><w:body>{content}</w:body></w:document>')
        with zipfile.ZipFile(path, 'w') as z:
            z.writestr('word/document.xml', xml)
        return path

    def test_blank_pasport_no_and_literal_table_marker(self):
        path = self.doc(tagged('PASSPORT_NO') + '<w:p><w:r><w:t>{{VIK_TABLE}}</w:t></w:r></w:p>')
        tokens, controls = inspect_word_fields(path)
        self.assertIn('NUMBER', controls)
        self.assertIn('VIK_TABLE', tokens)

    def test_observed_vik_field_aliases(self):
        from production_app.word_controls import normalized_tag
        self.assertEqual(normalized_tag('PASSPORT_NO'), 'NUMBER')
        self.assertEqual(normalized_tag('DOGOVOR_NO'), 'CONTRACT')
        self.assertEqual(normalized_tag('KMD_CODE'), 'CIPHER')
        self.assertEqual(normalized_tag('STEEL_MARK'), 'STEEL')
        self.assertEqual(normalized_tag('VIK_DATE'), 'VIK_DATE')

    def test_historical_vik_without_mark_table_is_accepted_by_contract(self):
        from production_app.templates import WORD_REQUIRED
        self.assertNotIn('VIK_TABLE', WORD_REQUIRED['vik'])
        self.assertIn('EVIDENCE', WORD_REQUIRED['vik'])
        self.assertIn('RESPONSIBLE', WORD_REQUIRED['vik'])

    def test_filled_historical_number_is_refused(self):
        with self.assertRaisesRegex(Stop, 'историческими данными'):
            inspect_word_fields(self.doc(tagged('PASSPORT_NO', '18/1/2026')))

    def test_native_table_control_does_not_replace_table_token(self):
        _, controls = inspect_word_fields(self.doc(tagged('VIK_TABLE')))
        self.assertNotIn('VIK_TABLE', controls)

    def test_empty_control_is_bound(self):
        control = FakeControl('PASSPORT_NO')
        doc = type('Document', (), {'ContentControls': FakeCollection(control)})()
        self.assertEqual(fill_word_controls(doc, {'NUMBER': '19/1/2026'}), {'NUMBER'})
        self.assertEqual(control.Range.Text, '19/1/2026')

    def test_filled_control_is_not_overwritten(self):
        control = FakeControl('PASSPORT_NO', '18/1/2026')
        doc = type('Document', (), {'ContentControls': FakeCollection(control)})()
        with self.assertRaisesRegex(Stop, 'предыдущего паспорта'):
            fill_word_controls(doc, {'NUMBER': '19/1/2026'})
        self.assertEqual(control.Range.Text, '18/1/2026')

    def test_locked_control_is_refused(self):
        control = FakeControl('PASSPORT_NO', locked=True)
        doc = type('Document', (), {'ContentControls': FakeCollection(control)})()
        with self.assertRaisesRegex(Stop, 'заблокировано'):
            fill_word_controls(doc, {'NUMBER': '19/1/2026'})


if __name__ == '__main__':
    unittest.main(verbosity=2)
