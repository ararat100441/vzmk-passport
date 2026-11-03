"""After-commit checks on SHA-pinned immutable artifacts, with real NTFS links.

These tests require Windows for the native hardlink and pywin32 path checks.
They do not assert that the frozen EXE or Microsoft Office has been exercised.
"""
import os
from pathlib import Path
from types import SimpleNamespace
from unittest import TestCase
from unittest.mock import MagicMock, patch

from openpyxl import Workbook

import test_codex50 as fixture
from production_app import core, storage
from production_app.shared_uzk import read_shared_uzk


class PostCommitArtifacts(TestCase):
    def setUp(self):
        self.fx = fixture.JournalAfterWrite(
            "test_linked_journal_readback_calls_native_verification"
        )
        self.fx.setUp()
        self.addCleanup(self.fx.doCleanups)
        self.fx.prepared_readback()

    def link(self, path, label):
        alias = self.fx.root / ("alias-" + label + Path(path).suffix)
        os.link(path, alias)
        self.addCleanup(alias.unlink, missing_ok=True)
        with self.assertRaisesRegex(storage.Stop, "Hardlink"):
            storage.safe(path)
        return alias

    def readback_with_office_boundary_mock(self):
        registry = SimpleNamespace(
            projects={"TEST": [dict(number="11/2026", row=2)]},
            sheet=1,
            columns={"number": 1},
        )
        app = MagicMock()
        app.Workbooks.Open.return_value.Worksheets.return_value.Cells.return_value.Value2 = "11/2026"
        with (
            patch("production_app.coordinated.native_readback") as native,
            patch("production_app.core.excel_session") as excel,
            patch("production_app.core.read_registry", return_value=registry),
            patch(
                "production_app.print_pdf.verify_package_print",
                return_value={
                    "filename": self.fx.pdf.name,
                    "pages": 1,
                    "sha256": storage.sha(self.fx.pdf),
                },
            ),
            patch("production_app.print_pdf.verify_print_pdf"),
            patch("production_app.core.office.verify_documents"),
        ):
            excel.return_value.__enter__.return_value = app
            self.assertEqual(core.readback(self.fx.journal)["status"], "PASS")
            native.assert_called_once()

    def test_verified_backup_hardlinks_do_not_interrupt_readback(self):
        for item in self.fx.state["participants"]:
            self.link(item["backup"], "backup-" + item["role"])
        context = core.post_commit_context(self.fx.journal)
        self.assertEqual(context[0]["phase"], "NEEDS_READBACK")
        self.readback_with_office_boundary_mock()

    def test_verified_plan_issue_and_pdf_hardlinks_do_not_interrupt_readback(self):
        self.link(self.fx.run / "plan.json", "plan")
        self.link(self.fx.target / "ISSUE.json", "issue")
        self.link(self.fx.pdf, "pdf")
        self.readback_with_office_boundary_mock()

    def test_corrupted_backup_remains_blocking_even_with_second_link(self):
        backup = Path(self.fx.state["participants"][0]["backup"])
        self.link(backup, "backup-registry")
        backup.write_bytes(b"corrupted verified backup")
        with self.assertRaisesRegex(storage.Stop, "Резервная копия повреждена"):
            core.post_commit_context(self.fx.journal)

    def test_reparse_metadata_on_backup_remains_blocking(self):
        backup = Path(self.fx.state["participants"][0]["backup"])
        self.link(backup, "backup-registry")
        actual = storage.windows_file_identity

        def altered(path):
            info = actual(path)
            if Path(path) == backup:
                info["attributes"] |= 0x400
            return info

        with patch(
            "production_app.storage.windows_file_identity",
            side_effect=altered,
        ):
            with self.assertRaisesRegex(storage.Stop, "reparse"):
                core.post_commit_context(self.fx.journal)

    def test_readonly_uzk_backup_parser_tolerates_link_only_in_verified_context(self):
        backup = self.fx.root / "uzk-native-backup.xlsm"
        book = Workbook()
        sheet = book.active
        sheet.title = "УЗК"
        sheet.append(["Заказ", "Марка элемента", "Кол-во (шт)"])
        sheet.append(["TEST", "К1-3", 2])
        book.save(backup)
        book.close()
        self.link(backup, "uzk-native-backup")
        with self.assertRaisesRegex(storage.Stop, "Hardlink"):
            read_shared_uzk(backup, project="TEST", marks={"К1-3": 1})
        actual = read_shared_uzk(
            backup, project="TEST", marks={"К1-3": 1}, post_commit=True
        )
        self.assertEqual(actual[("test", "к1-3")]["remaining"], 2)

    def test_complete_after_verified_backup_hardlink_preserves_commit_safety(self):
        for item in self.fx.state["participants"]:
            self.link(item["backup"], "backup-" + item["role"])
        with patch("production_app.core.supervisor_proof"):
            result = core.complete(self.fx.journal, self.fx.root / "result.json")
        self.assertEqual(result["state"], "ГОТОВО")
        self.assertEqual(core.journal_read(self.fx.journal)["phase"], "COMPLETE")
        self.assertEqual(core.pending(self.fx.registry), [])


if __name__ == "__main__":
    import unittest

    unittest.main()
