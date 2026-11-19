"""Classic screen stays available over the final document workers."""
import tempfile,unittest,tkinter as tk
from decimal import Decimal
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch
from production_app.gui import App
from production_app.storage import save

class ClassicInterface(unittest.TestCase):
    def setUp(self):
        self.temp=tempfile.TemporaryDirectory();self.folder=Path(self.temp.name)
        save(self.folder/'rc3-marked-settings.json',{'registry':'COPY.xlsx','profiles':{'Alpha':{}}})
        self.first_inventory={'sources':{},'missing':['appendix'],'passport_root':'Alpha/Паспорт'}
        self.patches=[patch('production_app.gui.local_root',return_value=self.folder),
          patch('production_app.gui.model.read_registry',return_value=SimpleNamespace(projects={'Alpha':[{'number':'1/2026'}]})),
          patch('production_app.gui.model.select',return_value=SimpleNamespace(first_number='1/2026',last_number='10/2026',next_number='11/2026')),
          patch('production_app.gui.model.approved_legacy_aliases',return_value=[]),
          patch('production_app.first_passport.discover_first_passport_for_config',return_value=self.first_inventory),
          patch('production_app.gui.messagebox.showerror')]
        for p in self.patches:p.start()
        self.root=tk.Tk();self.root.withdraw();self.app=App(self.root)
    def tearDown(self):
        self.root.destroy()
        for p in reversed(self.patches):p.stop()
        self.temp.cleanup()
    def test_rc3_stage_buttons_and_historical_number_selection_remain(self):
        self.assertEqual([b['text'] for b in self.app.buttons],['Сформировать пакет'])
        self.assertEqual(self.app.number.get(),'11/2026')
        self.app.new.set(False);self.app.select_project();self.assertEqual(self.app.number.get(),'1/2026')
    def test_profile_does_not_overwrite_the_installed_daily_settings(self):
        self.assertEqual(self.app.settings_path(),self.folder/'rc3-marked-settings.json')
        self.assertFalse((self.folder/'rc3-ux-settings.json').exists())
        self.assertFalse((self.folder/'rc3-autofirst-settings.json').exists())
        self.assertFalse((self.folder/'rc3-ui-settings.json').exists())
        self.assertFalse((self.folder/'settings.json').exists())
    def test_actual_vom_summary_needs_no_control_workbook_or_results(self):
        v=SimpleNamespace(rows=[None]*104,quantity=Decimal(288),mass=Decimal('15822.14'),area=Decimal('1021.644'))
        with patch('production_app.gui.model.visible_vom',return_value=v):self.assertIs(self.app.summarize_vom(),v)
        text=self.app.log.get('1.0','end');self.assertIn('104 марок / 288 изделий',text);self.assertIn('15822.14',text);self.assertIn('1021.644',text)
        self.assertEqual(self.app.facts,{})
    def test_unconfirmed_controls_never_start_production_prepare(self):
        self.app.cfg.update(documents='forms',output='out');self.app.vom.set('real.xlsx')
        with patch.object(self.app,'run') as run:
            self.app.check()
            run.assert_not_called()
            self.assertIsNone(self.app.plan)
        with patch.object(self.app,'run') as run:
            self.app.check_official();run.assert_not_called()
    def test_project_selection_finds_first_passport_without_loading_old_results(self):
        self.assertIs(self.app.first_reference,self.first_inventory)
        self.assertTrue(self.app.cfg['auto_first_passport'])
        self.assertEqual(self.app.facts,{})
        self.assertIn('Автоматически найден первый паспорт',self.app.log.get('1.0','end'))
    def test_verify_requires_private_pdf_before_user_confirmation(self):
        self.app.plan='frozen-plan';pdf=self.folder/'draft.pdf';pdf.write_bytes(b'SYNTHETIC FIXTURE')
        with patch.object(self.app,'review_preview') as review:
            self.app.verified({'preview_pdf':str(pdf)})
            review.assert_called_once()
            self.assertEqual(self.app.buttons[0]['text'],'Просмотреть и подтвердить')
    def test_open_prepared_documents_opens_review_copy_with_pdf(self):
        job=self.folder/'job';job.mkdir();preview=job/'preview';preview.mkdir();save(job/'preview.json',{'pdf':str(preview/'all.pdf')})
        self.app.plan=str(job/'plan.json')
        with patch('production_app.gui.os.startfile') as opened:self.app.open_documents();opened.assert_called_once_with(preview)

if __name__=='__main__':unittest.main()
