import unittest
from unittest.mock import patch
from pathlib import Path
from production_app import core
from production_app.storage import read,save,sha,Stop
import test_print_release as release_fixture

class PreviewBeforeApproval(unittest.TestCase):
    fake_export=staticmethod(release_fixture.ReleasePrintIntegration.fake_export)
    tearDown=release_fixture.ReleasePrintIntegration.tearDown
    def setUp(self):
        release_fixture.ReleasePrintIntegration.setUp(self)
        self.plan['facts']={'responsible':'Operator'}
        p=read(self.run/'prepared.json')|{'plan_seal':self.plan['seal'],'output_fingerprints':self.sources_hash};save(self.run/'prepared.json',p)
        v=read(self.run/'verified.json')|{'plan_seal':self.plan['seal']};save(self.run/'verified.json',v)
        (self.run/'confirmation.json').unlink()
    def preview(self):
        with patch('production_app.core.load_plan',return_value=self.plan),patch('production_app.print_pdf._export_word',side_effect=self.fake_export),patch('production_app.print_pdf._export_excel',side_effect=self.fake_export):return core.preview(self.plan_path)
    def test_preview_is_private_without_authorization_or_registry_write(self):
        r=self.preview();self.assertEqual(r['pages'],6);self.assertTrue(Path(r['pdf']).is_relative_to(self.run))
        self.assertFalse((self.run/'confirmation.json').exists());self.assertFalse(self.target.exists());self.assertEqual(sha(self.registry),self.registry_hash)
        self.assertEqual({p:sha(p) for p in self.sources_hash},self.sources_hash)
    def test_changed_preview_refuses_confirmation(self):
        r=self.preview();Path(r['pdf']).write_bytes(b'CHANGED')
        with patch('production_app.core.load_plan',return_value=self.plan),self.assertRaises(Stop):core.confirm(self.plan_path,'Operator',True)
        self.assertFalse((self.run/'confirmation.json').exists())
    def test_review_is_not_assumed_from_preview_creation(self):
        self.preview()
        with patch('production_app.core.load_plan',return_value=self.plan),self.assertRaises(Stop):core.confirm(self.plan_path,'Operator',False)
        self.assertFalse((self.run/'confirmation.json').exists())

if __name__=='__main__':unittest.main()
