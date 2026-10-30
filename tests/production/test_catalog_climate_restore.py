"""CODEX29: nested source settings and climate visibility cannot erase inputs."""
import tempfile,unittest,tkinter as tk
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch
from tkinter import ttk
from openpyxl import Workbook
from production_app.first_passport import archive_root_for_project,discover_first_passport_for_config
from production_app.path_config import normalize_project_roots
from production_app.akz_ui import climate_variants,display_variants
from production_app.gui import App
from production_app.storage import Stop,sha


class NestedProjectRoots(unittest.TestCase):
    def setUp(self):
        self.tmp=tempfile.TemporaryDirectory();self.addCleanup(self.tmp.cleanup)
        self.root=Path(self.tmp.name)
        self.project=self.root/'Гараж казахстан';self.passports=self.project/'Паспорт'
        self.leaf=self.passports/'2 от 15.09.2026';self.leaf.mkdir(parents=True)
    def test_dated_passport_project_and_passport_folder_resolve_same_archive(self):
        for path in (self.leaf,self.passports,self.project,self.root):
            self.assertEqual(archive_root_for_project(path,'Гараж Казахстан'),self.root)
    def test_other_project_and_unrelated_folder_do_not_fall_back(self):
        self.assertEqual(archive_root_for_project(self.leaf,'ТЭМПО 1'),self.leaf)
        unrelated=self.root/'Unrelated';unrelated.mkdir()
        self.assertEqual(archive_root_for_project(unrelated,'Гараж Казахстан'),unrelated)
    def test_normalization_is_in_memory_and_retains_registry_and_forms(self):
        cfg={'first_passport_documents':str(self.leaf),'output':str(self.leaf),
             'registry':'registry.xlsx','documents':'Templates','training':False}
        result=normalize_project_roots(cfg,'Гараж Казахстан')
        self.assertEqual(result['first_passport_documents'],str(self.root))
        self.assertEqual(result['output'],str(self.root))
        self.assertEqual(result['registry'],cfg['registry'])
        self.assertEqual(result['documents'],cfg['documents'])
        self.assertEqual(cfg['output'],str(self.leaf))
    def test_training_output_never_changes_to_original_archive(self):
        dev=self.root/'DEV';dev.mkdir()
        cfg={'first_passport_documents':str(self.leaf),'output':str(dev),'training':True}
        result=normalize_project_roots(cfg,'Гараж Казахстан')
        self.assertEqual(result['first_passport_documents'],str(self.root))
        self.assertEqual(result['output'],str(dev))
    def test_explicit_unrelated_archive_is_not_overridden(self):
        unrelated=self.root/'Other';unrelated.mkdir()
        cfg={'first_passport_documents':str(unrelated),'output':str(self.leaf)}
        self.assertEqual(normalize_project_roots(cfg,'Гараж Казахстан'),cfg)
    def test_discovery_uses_recognized_archive_not_second_passport_contents(self):
        project=SimpleNamespace(name='Гараж Казахстан')
        with patch('production_app.first_passport.discover_first_passport',return_value={}) as scan:
            discover_first_passport_for_config({'first_passport_documents':str(self.leaf)},project)
        self.assertEqual(scan.call_args.args,(self.root,project))
    def test_duplicate_passport_folders_are_not_guessed(self):
        (self.project/' ПАСПОРТ ').mkdir()
        with self.assertRaisesRegex(Stop,'единственная'):
            archive_root_for_project(self.leaf,'Гараж Казахстан')


class ClimateVisibility(unittest.TestCase):
    def test_sole_act_shows_climate_even_without_appendix_or_assignments(self):
        profile={'templates':{'akz0':{'path':'APPROVED'}}}
        rows=[{'mark':'M','akz':None}]
        self.assertEqual(climate_variants(profile),[0])
        self.assertEqual(display_variants(rows,profile),[0])
        self.assertIsNone(rows[0]['akz'])
    def test_partial_assignments_do_not_invent_another_variant(self):
        profile={'templates':{'akz0':{},'akz3':{}}}
        self.assertEqual(display_variants([{'akz':None},{'akz':0}],profile),[0])
        self.assertEqual(display_variants([{'akz':None}],profile),[])


class NativeFactsScreen(unittest.TestCase):
    def setUp(self):
        self.tmp=tempfile.TemporaryDirectory();self.addCleanup(self.tmp.cleanup)
        self.path=Path(self.tmp.name)/'akz0.xlsx'
        book=Workbook();book.active['A1']='{{AKZ_THICKNESS}}';book.save(self.path);book.close()
        self.root=tk.Tk();self.root.withdraw();self.addCleanup(self.root.destroy)
        self.app=App.__new__(App);self.app.root=self.root;self.app.active=None
        self.app.cfg={'training':True,'profiles':{}}
        self.app.project=tk.StringVar(self.root,value='Гараж Казахстан')
        self.app.number=tk.StringVar(self.root,value='11/2026')
        self.app.vom=tk.StringVar(self.root,value='SYNTHETIC.xlsx')
        self.app.facts={};self.app.first_reference=None;self.app.buttons=[]
        self.messages=[];self.app.write=self.messages.append;self.app.error=lambda e:self.fail(str(e))
        profile={'templates':{'akz0':{'path':str(self.path),'sha256':sha(self.path)}}}
        patches=[patch('production_app.gui.model.visible_vom',return_value=SimpleNamespace(rows=[('M',)])),
                 patch('production_app.shared_uzk.find_shared_uzk',return_value='SYNTHETIC.xlsm'),
                 patch('production_app.shared_uzk.planned_control',return_value={'M':{'uzk_quantity':8}}),
                 patch('production_app.gui.model.profile',side_effect=Stop('Каталог недоступен')),
                 patch('production_app.prepared_forms.prepared_project',return_value=profile)]
        for p in patches:p.start();self.addCleanup(p.stop)
    def descendants(self,widget):
        for child in widget.winfo_children():
            yield child;yield from self.descendants(child)
    def open_screen(self):
        self.app.fact_dialog();self.root.update_idletasks()
        return next(w for w in self.root.winfo_children() if isinstance(w,tk.Toplevel))
    def test_catalog_failure_does_not_hide_approved_one_layer_climate(self):
        dialog=self.open_screen()
        frames=[w for w in self.descendants(dialog) if isinstance(w,ttk.LabelFrame) and w['text'].startswith('АКЗ №')]
        self.assertEqual(len(frames),1)
        self.assertIn('1 слой',frames[0]['text'])
        self.assertEqual(frames[0].winfo_manager(),'pack')
        entries=[w for w in self.descendants(frames[0]) if isinstance(w,ttk.Entry) and int(w.grid_info()['row'])<=6]
        self.assertEqual(len(entries),10)
        grid=next(w for w in self.descendants(dialog) if isinstance(w,ttk.Treeview))
        self.assertEqual(grid.item(grid.get_children()[0],'values'),('M','','','8',''))
    def test_partial_last_climate_edit_survives_back_and_reopen(self):
        dialog=self.open_screen()
        frame=next(w for w in self.descendants(dialog) if isinstance(w,ttk.LabelFrame) and w['text'].startswith('АКЗ №'))
        entry=next(w for w in self.descendants(frame) if isinstance(w,ttk.Entry) and w.grid_info().get('row')==6 and w.grid_info().get('column')==2)
        entry.insert(0,'99')
        next(w for w in self.descendants(dialog) if isinstance(w,ttk.Button) and w['text']=='← Назад').invoke()
        self.assertEqual(self.app.facts['akz_drafts']['0']['layer_intervals'][0]['max'],'99')
        self.assertEqual(self.app.facts['marks'][0]['akz'],'')
        self.assertFalse(self.app.facts['confirmed'])
        reopened=self.open_screen()
        frame=next(w for w in self.descendants(reopened) if isinstance(w,ttk.LabelFrame) and w['text'].startswith('АКЗ №'))
        entry=next(w for w in self.descendants(frame) if isinstance(w,ttk.Entry) and w.grid_info().get('row')==6 and w.grid_info().get('column')==2)
        self.assertEqual(entry.get(),'99')


if __name__=='__main__':unittest.main()
