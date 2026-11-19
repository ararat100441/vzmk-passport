import tempfile,unittest,tkinter as tk
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch
from production_app.daily_gui import DailyApp,explain_error
from production_app.storage import save

class DailyGUI(unittest.TestCase):
    def setUp(self):
        self.temp=tempfile.TemporaryDirectory();self.folder=Path(self.temp.name)
        save(self.folder/'settings.json',{'registry':'test-copy.xlsx','training':True,'profiles':{
          'Alpha':{'preferences':{'doc_fields':{'QUALITY_MANUFACTURER':'Alpha factory'}}},
          'Beta':{'preferences':{'doc_fields':{'QUALITY_MANUFACTURER':'Beta factory'}}}}})
        self.patches=[patch('production_app.daily_gui.local_root',return_value=self.folder),
          patch('production_app.daily_gui.model.read_registry',return_value=SimpleNamespace(projects={'Alpha':[],'Beta':[]})),
          patch('production_app.daily_gui.model.select',side_effect=lambda r,n:SimpleNamespace(next_number='2/2026' if n=='Alpha' else '9/1/2026')),
          patch('production_app.daily_gui.messagebox.showerror')]
        for p in self.patches:p.start()
        self.root=tk.Tk();self.root.withdraw();self.app=DailyApp(self.root)
    def tearDown(self):
        self.root.destroy()
        for p in reversed(self.patches):p.stop()
        self.temp.cleanup()
    def test_one_primary_action_and_no_exposed_engine_buttons(self):
        self.assertEqual(len(self.app.buttons),1);self.assertEqual(self.app.primary['text'],'Сформировать пакет')
        self.assertFalse(self.app.reviewed.get());self.assertFalse(self.app.approved.get())
    def test_project_switch_never_keeps_other_shipment_facts(self):
        a=self.app;a.project.set('Alpha');a.select_project();self.assertEqual(a.number.get(),'2/2026')
        a.draft={'evidence':['old protocol'],'climates':[{'average':123}]};a.approved.set(True)
        a.project.set('Beta');a.select_project();self.assertEqual(a.number.get(),'9/1/2026');self.assertEqual(a.draft,{})
        self.assertFalse(a.approved.get());self.assertEqual(a.current_preferences()['doc_fields']['QUALITY_MANUFACTURER'],'Beta factory')
    def test_release_requires_both_open_pdf_and_explicit_review(self):
        a=self.app;a.plan='private-plan';a.facts={'responsible':'Operator'}
        with patch.object(a,'run') as run:
            a.authorize();run.assert_not_called()
            a.opened=True;a.authorize();run.assert_not_called()
            a.reviewed.set(True);a.authorize();run.assert_not_called()
            a.approved.set(True);a.authorize();self.assertEqual(run.call_args.args[0],'confirm')
    def test_failed_project_selection_cannot_show_previous_number_or_shipment(self):
        a=self.app;a.project.set('Alpha');a.select_project();a.vom.set('old.xlsx');a.control_path.set('old-control.xlsx');a.summary.set('Old totals')
        with patch('production_app.daily_gui.model.select',side_effect=ValueError('Unsafe project name')):
            a.project.set('Beta');a.select_project()
        self.assertEqual(a.number.get(),'');self.assertEqual(a.vom.get(),'');self.assertEqual(a.control_path.get(),'')
        self.assertIn('Beta',a.summary.get());self.assertNotIn('Old totals',a.summary.get())
    def test_other_project_cannot_open_previous_projects_completed_pdf(self):
        a=self.app;a.result_project='Alpha';a.final_pdf='alpha.pdf';a.final_folder='alpha-folder'
        a.project.set('Beta');a.select_project()
        self.assertIsNone(a.final_pdf);self.assertIsNone(a.final_folder)
    def test_refresh_same_project_keeps_its_completed_package_accessible(self):
        a=self.app;a.project.set('Alpha');a.result_project='Alpha';a.final_pdf='alpha.pdf';a.final_folder='alpha-folder'
        a.select_project()
        self.assertEqual(a.final_pdf,'alpha.pdf');self.assertEqual(a.final_folder,'alpha-folder')
    def test_worker_cannot_lose_frozen_plan_due_to_input_trace(self):
        a=self.app;a.plan='frozen-plan';a.active=object();a.day.set('2026-10-25');self.assertEqual(a.plan,'frozen-plan');a.active=None
    def test_changed_operator_invalidates_previous_review_and_approval(self):
        a=self.app;a.plan='old-plan';a.reviewed.set(True);a.approved.set(True);a.operator.set('Other operator')
        self.assertIsNone(a.plan);self.assertFalse(a.approved.get());self.assertFalse(a.reviewed.get())
    def test_cancel_after_preparation_never_starts_release_or_commit(self):
        a=self.app;a.cancel_requested=True
        with patch.object(a,'run') as run:
            a.after_stage('prepare',lambda r:a.run('release',{},None),{'status':'PASS'});run.assert_not_called()
        self.assertEqual(a.phase,'input')
    def test_cancel_during_commit_still_finishes_readback_instead_of_unknown_state(self):
        a=self.app;a.cancel_requested=True;calls=[]
        a.after_stage('commit',lambda r:calls.append('readback'),{'status':'PASS'})
        self.assertEqual(calls,['readback'])
    def test_climate_defaults_are_from_selected_project_and_never_include_means(self):
        a=self.app;a.cfg['training']=False
        for name in ('Alpha','Beta'):a.cfg['profiles'][name]['preferences']['assignments']={'M1':{'akz':0}}
        v=SimpleNamespace(rows=[('M1',)])
        def defaults(name):return {'variants':{'0':{'climate':{'air':'21' if name=='Alpha' else '26'},'layers':[{'layer':1,'min':'80','max':'130'}]}}}
        with patch('production_app.daily_gui.akz_presets.last_for_project',side_effect=defaults):
            a.project.set('Alpha');alpha=a.fresh_draft(v);a.project.set('Beta');beta=a.fresh_draft(v)
        self.assertEqual(alpha['climates'][0]['air'],'21');self.assertEqual(beta['climates'][0]['air'],'26')
        self.assertNotIn('dry_averages',beta['climates'][0]);self.assertNotIn('layer_details',beta['climates'][0]);self.assertNotIn('evidence',beta)
    def test_open_folder_during_review_never_opens_previous_passport(self):
        a=self.app;current=self.folder/'current-review';current.mkdir();previous=self.folder/'previous-passport';previous.mkdir()
        a.phase='review';a.preview_path=str(current/'new.pdf');a.final_folder=str(previous)
        with patch('production_app.daily_gui.os.startfile') as opened:
            a.open_folder();opened.assert_called_once_with(str(current))
    def test_missing_fields_use_the_same_labels_as_the_input_form(self):
        self.assertIn('Дата ВИК',explain_error('Не подтверждена дата: vik_date'))
        self.assertIn('Партия отвердителя',explain_error('Не подтверждено hardener_batch слоя 2'))
        self.assertIn('Частота, МГц',explain_error('Не заполнено поле frequency_mhz в УЗК, строка 3'))
    def test_error_translation_never_changes_a_source_path(self):
        text='Файл занят: C:/work/date/1004/material.xlsx\nЖурнал: C:/logs/date'
        self.assertEqual(explain_error(text),text)

if __name__=='__main__':unittest.main()
