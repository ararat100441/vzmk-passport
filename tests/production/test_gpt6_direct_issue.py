"""Direct issue without preliminary PDF; no false document-review attestation.

These tests cover GUI routing and authorization. Native Windows Office issuance
must still be tested on isolated copies before releasing a new EXE.
"""
from pathlib import Path
from tempfile import TemporaryDirectory
from unittest import TestCase
from unittest.mock import MagicMock, Mock, patch

from production_app import core, gui
from production_app.owner_control import require_declaration
from production_app.storage import Stop, read, save


class DirectIssueGui(TestCase):
    def app(self):
        instance = object.__new__(gui.App)
        instance.active = None
        instance.plan = "P/plan.json"
        instance.preview_pdf = None
        instance.buttons = [MagicMock()]
        instance.write = Mock()
        instance.error = Mock()
        instance.issue = Mock()
        return instance

    def context(self):
        d = {
            "seal": "PLAN_SEAL",
            "project": "Гараж Казахстан",
            "number": "11/2026",
            "date": "2026-11-03",
            "marks": 3,
            "quantity": 3,
            "mass_kg": "842.92",
            "area_m2": "26.877",
            "registry": "REGISTRY.xlsx",
            "target": "OUTPUT",
            "accounting": {
                "catalog": "CATALOG.xlsx",
                "uzk": "УЗК.xlsm",
                "controls": {"К1-3": 1},
                "stock_warnings": [],
            },
        }
        return d, {"plan_seal": d["seal"], "preview_pdf": None,
                   "documents": [{"file": "quality.docx", "status": "PASS"}]}

    def test_prepared_skips_preliminary_pdf_creation(self):
        a = self.app()
        a.run = Mock()
        a.prepared({"files": 6})
        self.assertEqual(a.run.call_args.args[0], "verify")
        self.assertEqual(a.run.call_args.args[1],
                         {"path": a.plan, "preview": False})

    def test_direct_confirmation_never_opens_preview_and_invokes_issue_once(self):
        a = self.app()
        d, proof = self.context()
        with patch("production_app.gui.read", side_effect=[d, proof]), \
             patch("production_app.gui.messagebox.askyesno", return_value=True) as confirm, \
             patch("production_app.gui.open_external") as opened:
            a.verified({"preview_pdf": None, "ready_to_confirm": True})
        opened.assert_not_called()
        a.issue.assert_called_once()
        confirm.assert_called_once()
        self.assertIn("БЕЗ предварительного PDF", confirm.call_args.args[1])
        self.assertIn("REGISTRY.xlsx", confirm.call_args.args[1])

    def test_cancel_does_not_issue_or_write(self):
        a = self.app()
        d, proof = self.context()
        with patch("production_app.gui.read", side_effect=[d, proof]), \
             patch("production_app.gui.messagebox.askyesno", return_value=False):
            a.verified({"preview_pdf": None, "ready_to_confirm": True})
        a.issue.assert_not_called()
        a.write.assert_called_once()

    def test_issue_seals_direct_mode_without_claiming_pdf_review(self):
        a = self.app()
        a.facts = {"responsible": "Operator"}
        a.run = Mock()
        a.issue()
        task, payload = a.run.call_args.args[:2]
        self.assertEqual(task, "confirm")
        self.assertFalse(payload["reviewed"])
        self.assertTrue(payload["direct_issue"])
        self.assertTrue(payload["approved"])

    def test_previously_prepared_plan_must_not_restart_check(self):
        a = self.app()
        a.run = Mock()
        a.review_direct = Mock()
        a.check()
        a.review_direct.assert_called_once()
        a.run.assert_not_called()


class DirectIssueConfirmation(TestCase):
    def setUp(self):
        self.temp = TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.folder = Path(self.temp.name)
        self.plan_path = self.folder / "plan.json"
        self.plan = {
            "run": str(self.folder),
            "seal": "ONE_PLAN",
            "project": "Гараж Казахстан",
            "number": "11/2026",
            "facts": {"responsible": "Operator", "simple_ui": True,
                      "owner_confirmation_required": True},
        }
        self.verified = {
            "plan_seal": "ONE_PLAN",
            "preview_pdf": None,
            "fingerprints": {},
            "documents": [{"file": "quality.docx", "status": "PASS"}],
        }
        save(self.folder / "verified.json", self.verified)

    def confirm(self, **kwargs):
        with patch("production_app.core.load_plan", return_value=self.plan):
            return core.confirm(self.plan_path, "Operator", False, **kwargs)

    def test_explicit_no_preview_approval_is_recorded_truthfully(self):
        result = self.confirm(approved=True, current_control=True,
                              direct_issue=True)
        self.assertEqual(result["status"], "PASS")
        confirmation = read(self.folder / "confirmation.json")
        self.assertFalse(confirmation["reviewed"])
        self.assertTrue(confirmation["direct_issue_confirmed"])
        self.assertTrue(confirmation["approved"])
        self.assertEqual(require_declaration(self.plan, confirmation)["attested"], True)

    def test_direct_issue_without_user_approval_fails(self):
        with self.assertRaises(Stop):
            self.confirm(approved=False, current_control=True, direct_issue=True)
        self.assertFalse((self.folder / "confirmation.json").exists())

    def test_no_preview_without_direct_approval_fails(self):
        with self.assertRaises(Stop):
            self.confirm(approved=True, current_control=True, direct_issue=False)
        self.assertFalse((self.folder / "confirmation.json").exists())

    def test_direct_issue_disallows_stale_preview_state(self):
        self.verified["preview_pdf"] = "stale-preview.pdf"
        save(self.folder / "verified.json", self.verified)
        with self.assertRaisesRegex(Stop, "предварительным PDF"):
            self.confirm(approved=True, current_control=True, direct_issue=True)
        self.assertFalse((self.folder / "confirmation.json").exists())

    def test_direct_issue_requires_verified_documents(self):
        self.verified["documents"] = []
        save(self.folder / "verified.json", self.verified)
        with self.assertRaisesRegex(Stop, "Word и Excel"):
            self.confirm(approved=True, current_control=True, direct_issue=True)


class DirectIssueReadbackContract(TestCase):
    """The sealed post-commit context must recognize explicit no-PDF consent."""

    def setUp(self):
        import test_codex50 as fixture
        self.fixture = fixture.JournalAfterWrite(
            "test_linked_journal_readback_calls_native_verification"
        )
        self.fixture.setUp()
        self.addCleanup(self.fixture.doCleanups)
        self.fixture.prepared_readback()

    def set_mode(self, direct, reviewed):
        folder = self.fixture.run
        confirmation = read(folder / "confirmation.json")
        confirmation["reviewed"] = reviewed
        confirmation["direct_issue_confirmed"] = direct
        save(folder / "confirmation.json", confirmation)
        return confirmation

    def test_direct_approval_allows_verified_readback_context(self):
        self.set_mode(direct=True, reviewed=False)
        state, _, _, _ = core.post_commit_context(self.fixture.journal)
        self.assertEqual(state["phase"], "NEEDS_READBACK")

    def test_neither_preview_review_nor_direct_approval_blocks_readback(self):
        self.set_mode(direct=False, reviewed=False)
        with self.assertRaisesRegex(Stop, "подтверждения"):
            core.post_commit_context(self.fixture.journal)


if __name__ == "__main__":
    import unittest
    unittest.main()
