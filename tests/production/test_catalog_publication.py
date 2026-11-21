import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from openpyxl import Workbook, load_workbook

from production_app.project_catalog import catalog_name, create_catalog, resolve_catalog


class CatalogPublication(unittest.TestCase):
    def test_write_failure_does_not_block_retry(self):
        def full_disk(book, stream):
            stream.write(b'partial')
            raise OSError('disk full')

        with tempfile.TemporaryDirectory() as folder:
            with patch.object(Workbook, 'save', full_disk), self.assertRaises(OSError):
                create_catalog(folder, 'Test')
            self.assertIsNone(resolve_catalog(folder, 'Test'))
            self.assertEqual(list(Path(folder).iterdir()), [])
            result = create_catalog(folder, 'Test')
            book = load_workbook(result['path'])
            try:
                self.assertIn('Марки и отгрузки', book.sheetnames)
            finally:
                book.close()

    def test_competing_file_is_preserved(self):
        save = Workbook.save
        with tempfile.TemporaryDirectory() as folder:
            target = Path(folder) / catalog_name('Test')

            def competing_creator(book, stream):
                target.write_bytes(b'concurrent')
                save(book, stream)

            with patch.object(Workbook, 'save', competing_creator), self.assertRaises(FileExistsError):
                create_catalog(folder, 'Test')
            self.assertEqual(target.read_bytes(), b'concurrent')
            self.assertEqual(list(Path(folder).iterdir()), [target])
