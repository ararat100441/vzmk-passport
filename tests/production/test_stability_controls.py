from decimal import Decimal
from pathlib import Path
from types import SimpleNamespace
import tempfile
import unittest
from unittest.mock import patch

from openpyxl import Workbook

from production_app import core
from production_app.amendment import control_plan, retained_design
from production_app.automatic_acts import calculate_design_characteristics
from production_app.owner_control import STATEMENT, declaration, require_declaration
from production_app.project_visibility import closed_projects, save_settings, set_closed
from production_app.shared_uzk import HEADERS, planned_control, verify_planned_quantities
from production_app.shipment_input import manual_row, parse_tsv
from production_app.storage import Stop, read, save, sha


class SourceSeparation(unittest.TestCase):
    def setUp(self):
        folder=tempfile.TemporaryDirectory()
        self.addCleanup(folder.cleanup)
        self.path=Path(folder.name)/'UZK.xlsm'
        book=Workbook();sheet=book.active;sheet.title='УЗК'
        sheet.append(HEADERS);sheet.append(('P','AUTO',5));sheet.append(('P','MIX',1))
        book.save(self.path);book.close()

    def test_manual_control_is_bounded_and_only_absent_mark_is_editable(self):
        v=SimpleNamespace(rows=(('AUTO','Beam',Decimal(2)),('FREE','Beam',Decimal(2))))
        plan=planned_control(self.path,'P',v)
        for mark,total,asked in [('AUTO',1,0),('AUTO',3,1),('FREE',3,3),('FREE',1,0)]:
            with self.subTest(mark=mark,total=total,asked=asked),self.assertRaises(Stop):
                rows={m:{'mark':m,'uzk_quantity':p['uzk_quantity']} for m,p in plan.items()}
                rows[mark].update(uzk_quantity=total,uzk_manual_requested=asked)
                verify_planned_quantities({'marks':list(rows.values())},plan,accept_manual=True)
        facts={'marks':[{'mark':'AUTO','uzk_quantity':2},{'mark':'FREE','uzk_quantity':2,
            'uzk_manual_requested':2,'profile':'UNKNOWN','uzk_length_mm':'75','uzk_thickness_mm':'8'}]}
        verify_planned_quantities(facts,plan,accept_manual=True)
        row=facts['marks'][1]
        self.assertEqual((row['uzk_registry_quantity'],row['uzk_manual_quantity']),(0,2))
        self.assertEqual(plan['FREE']['uzk_quantity'],0)
        design=calculate_design_characteristics(SimpleNamespace(rows=(v.rows[1],)),{'marks':[row]},None)
        self.assertEqual([(d['length_mm'],d['thickness_mm']) for d in design],[('75','8')])

    def test_previous_manual_control_survives_appearance_in_xlsm(self):
        old={'marks':[{'mark':'MIX','uzk_quantity':1,'uzk_manual_quantity':1,'uzk_registry_quantity':0}],
             'uzk_design':[{'mark':'MIX','quantity':1,'control_quantity':1,'control_source':'MANUAL',
                           'length_mm':'75','geometry_reference':{'source':'MANUAL'}}]}
        empty=SimpleNamespace(rows=())
        self.assertEqual(control_plan(self.path,'P',empty,old,{'MIX':1},{},allow_manual=True),({},{},{}))
        delta=SimpleNamespace(rows=(('MIX','Beam',Decimal(1)),))
        plan,debited,refunds=control_plan(self.path,'P',delta,old,{'MIX':1},{'MIX':2},allow_manual=True)
        self.assertEqual(debited['MIX']['uzk_quantity'],1)
        self.assertFalse(refunds)
        row={'mark':'MIX','uzk_quantity':2}
        verify_planned_quantities({'marks':[row]},plan,accept_manual=True)
        self.assertEqual((row['uzk_registry_quantity'],row['uzk_manual_quantity']),(1,1))
        self.assertEqual(retained_design(old,plan),old['uzk_design'])
        mixed={'marks':[row]}
        self.assertEqual(control_plan(self.path,'P',empty,mixed,{'MIX':2},{},allow_manual=True)[2],{'MIX':1})
        with self.assertRaisesRegex(Stop,'MIX'):
            control_plan(self.path,'P',empty,mixed,{'MIX':2},{'MIX':1},allow_manual=True)

    def test_confirmation_explicitly_covers_manual_control(self):
        verified=self.path.parent/'verified.json'
        save(verified,{'documents':[]})
        plan={'project':'P','number':'1/2026','seal':'DEV','run':str(verified.parent),
              'facts':{'marks':[{'uzk_registry_quantity':1,'uzk_manual_quantity':1}]}}
        checksum=sha(verified)
        owner=declaration(plan,checksum)
        confirmation={'current_control':owner,'approved':True,'reviewed':True,'verified_sha256':checksum}
        self.assertIn('включая дополнительный ручной УЗК',require_declaration(plan,confirmation)['statement'])
        with self.assertRaises(Stop):
            require_declaration(plan,{**confirmation,'current_control':{**owner,'statement':STATEMENT}})


class ShipmentAndProjects(unittest.TestCase):
    def test_manual_area_and_tsv_share_unknown_representation(self):
        self.assertEqual(manual_row('A','Beam','2','4',''),parse_tsv('A\tBeam\t2\t4\t8').rows[0])
        self.assertIsNone(manual_row('A','Beam','2','4','')[-1])
        for qty,mass,area in [('1.5','1',''),('1','NaN',''),('1','Infinity',''),('1','1e309',''),
                              ('1','1','-1'),('9999999999999999','1',''),('2','9e307','')]:
            with self.subTest(qty=qty,mass=mass,area=area),self.assertRaises(Stop):
                manual_row('A','Beam',qty,mass,area)

    def test_manual_row_without_area_survives_vom_editor_round_trip(self):
        import tkinter as tk
        from tkinter import ttk
        from production_app.shipment_dialog import open_shipment

        try:
            window=tk.Tk()
        except tk.TclError as exc:
            self.skipTest(str(exc))
        self.addCleanup(window.destroy)
        with tempfile.TemporaryDirectory() as folder:
            app=SimpleNamespace(root=window,active=None,fit=lambda w,x,y:w.geometry('900x500'),
                editable=lambda *args,**kwargs:None,vom=tk.StringVar(),vom_summary=tk.StringVar(),
                write=lambda _:None,error=lambda exc,**kw:self.fail(str(exc)))
            with patch('production_app.shipment_dialog.local_root',return_value=Path(folder)):
                editor=open_shipment(app)
                window.update()

                def find_button(root,label):
                    for child in root.winfo_children():
                        if isinstance(child,ttk.Button) and child['text']==label:
                            return child
                        found=find_button(child,label)
                        if found is not None:return found

                find_button(editor,'Добавить марку').invoke()
                form=next(child for child in editor.winfo_children() if isinstance(child,tk.Toplevel))
                fields=[child for child in form.winfo_children() if isinstance(child,ttk.Entry)]
                for field,value in zip(fields,('A','Beam','2','8')):
                    field.insert(0,value)
                find_button(form,'Добавить').invoke()
                find_button(editor,'ОК').invoke()
                self.assertEqual(app.shipment_rows,(manual_row('A','Beam','2','8',''),))
                self.assertIn('площадь не указана',app.vom_summary.get())
                self.assertTrue(Path(app.vom.get()).is_file())

    def test_visibility_does_not_change_books_and_survives_stale_settings_save(self):
        with tempfile.TemporaryDirectory() as folder:
            root=Path(folder);book=root/'registry.xlsx';book.write_bytes(b'DEV ledger')
            pin=sha(book);settings=root/'rc3-marked-settings.json'
            cfg={'registry':str(book),'particular_setting':17};save(settings,cfg)
            journal=root/'.vzmk-operations'/'run'/'journal.json'
            core.journal_write(journal,{'project':'P','number':'3/2026','phase':'NEEDS_READBACK'})
            with patch('production_app.model.read_registry',return_value=SimpleNamespace(projects={'P':[],'Q':[]})):
                with self.assertRaisesRegex(Stop,'3/2026'):
                    set_closed(settings,cfg,'P',closed=True)
                set_closed(settings,cfg,'Q',closed=True)
                self.assertEqual(closed_projects(read(settings)),('Q',))
                save_settings(settings,{**cfg,'particular_setting':23})
                self.assertEqual(closed_projects(read(settings)),('Q',))
                set_closed(settings,cfg,'Q',closed=False)
                self.assertEqual(closed_projects(read(settings)),())
            self.assertEqual(sha(book),pin)
            self.assertEqual(read(settings)['particular_setting'],23)
            self.assertEqual(core.journal_read(journal)['phase'],'NEEDS_READBACK')

    def test_corrupt_journal_reports_path_and_other_waiting_identity(self):
        with tempfile.TemporaryDirectory() as folder:
            root=Path(folder);book=root/'registry.xlsx';book.write_bytes(b'DEV ledger')
            broken=root/'.vzmk-operations'/'bad'/'journal.json'
            broken.parent.mkdir(parents=True);broken.write_bytes(b'{')
            waiting=broken.parent.parent/'valid'/'journal.json'
            core.journal_write(waiting,{'project':'P','number':'7/2026','phase':'NEEDS_READBACK'})
            with self.assertRaises(Stop) as raised:
                core.pending(book)
            self.assertTrue(all(x in str(raised.exception) for x in (str(broken),'P 7/2026','NEEDS_READBACK')))
            self.assertEqual(broken.read_bytes(),b'{')
