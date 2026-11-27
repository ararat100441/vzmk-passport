from decimal import Decimal
from pathlib import Path
from tempfile import TemporaryDirectory
from types import SimpleNamespace
import unittest
from unittest.mock import Mock, patch

from docx import Document
from openpyxl import Workbook

from production_app import office
from production_app.amendment import control_plan, retained_design
from production_app.automatic_acts import calculate_design_characteristics
from production_app.gui import App
from production_app.highlighted_reference import _normalize_design, create_uzk_copy
from production_app.native_acts import native_rows
from production_app.shared_uzk import HEADERS, current_design_quantity
from production_app.storage import Stop
from test_highlighted_reference import uzk_first


def historical_part(length, thickness, *, origin='REGISTRY', quantity=1, segments=(), manual=True):
    reference={'source':'MANUAL'} if manual else {'sheet':'Профили','row':683,'standard':'ГОСТ 8240-97'}
    return {'mark':'M','profile':'Швеллер 16П','quantity':quantity,'control_source':origin,
            'control_quantity':quantity,'length_mm':length,'thickness_mm':thickness,
            'frequency_mhz':'5','angle_deg':'70','segments':segments,'geometry_reference':reference}


class RepeatedAmendmentGeometry(unittest.TestCase):
    def test_new_registry_unit_uses_its_own_measurement_and_retains_prior_groups(self):
        previous=[historical_part('75','8',origin='MANUAL'),historical_part('155','6')]
        facts={'marks':[{'mark':'M','uzk_quantity':2,'uzk_registry_quantity':1,'uzk_manual_quantity':1}],
               'uzk_design':previous}
        shipment=SimpleNamespace(rows=(('M','DEV Beam',Decimal(1)),))
        with TemporaryDirectory() as folder:
            source=Path(folder)/'DEV.xlsx'
            book=Workbook();book.active.title='УЗК';book.active.append(HEADERS)
            book.active.append(('P','M',1));book.save(source);book.close()
            plan,_,refunds=control_plan(source,'P',shipment,facts,{'M':2},{'M':3},allow_manual=True)
        self.assertEqual(refunds,{})
        current={'mark':'M','profile':'Швеллер 16П','uzk_length_mm':'201','uzk_thickness_mm':'10'}
        current['uzk_quantity']=current_design_quantity(current,plan['M'])
        design=calculate_design_characteristics(shipment,{'marks':[current]},None)
        self.assertEqual([(r['length_mm'],r['thickness_mm'],r['quantity']) for r in design],[('201','10',1)])
        self.assertEqual(retained_design(facts,plan),previous)

    def test_part_removal_of_distinct_automatic_debits_with_manual_geometry_stops(self):
        facts={'marks':[{'mark':'M','uzk_quantity':2,'uzk_registry_quantity':2,'uzk_manual_quantity':0}],
               'uzk_design':[historical_part('155','6'),historical_part('201','10')]}
        with self.assertRaisesRegex(Stop,'M'):
            retained_design(facts,{'M':{'registry_retained':1}})

    def test_uniform_legacy_geometry_has_one_count_for_both_weld_segments(self):
        wall=historical_part('320','5',quantity=2,segments=('стенка',),manual=False)
        flange=historical_part('256','8.4',quantity=2,segments=('полки',),manual=False)
        for row in (wall,flange):
            row.pop('control_source');row.pop('control_quantity')
        facts={'marks':[{'mark':'M','uzk_quantity':2}],'uzk_design':[wall,flange]}
        result=retained_design(facts,{'M':{'registry_retained':1}})
        self.assertEqual([(p['length_mm'],p['control_quantity']) for p in result],[('160',1),('128',1)])

    def test_partial_removal_keeps_equivalent_catalog_geometry_without_merging_new_groups(self):
        old=[historical_part('160','5',segments=('стенка',),manual=False),
             historical_part('128','8.4',segments=('полки',),manual=False)]
        for row in old:row['geometry_reference'].update(path='C:/DEV/original.xlsx',sha256='a'*64)
        additional=[{**row,'geometry_reference':{**row['geometry_reference'],'path':'C:/DEV/rebound.xlsx'}}
                    for row in old]
        facts={'marks':[{'mark':'M','uzk_quantity':2,'uzk_registry_quantity':2,'uzk_manual_quantity':0}],
               'uzk_design':old+additional}
        kept=retained_design(facts,{'M':{'registry_retained':1}})
        self.assertEqual(kept,old)

    def test_new_amendment_does_not_seed_dimensions_from_issued_registry_measurement(self):
        historical={'marks':[{'mark':'M','uzk_quantity':1,'uzk_registry_quantity':1,
                              'uzk_manual_quantity':0,'uzk_length_mm':'155','uzk_thickness_mm':'6'}]}
        original=SimpleNamespace(rows=(),digest='DEV')
        application=SimpleNamespace(active=None,cfg={},new=SimpleNamespace(get=lambda:False),
                                    number=SimpleNamespace(get=lambda:'2/2026'),
                                    project=SimpleNamespace(get=lambda:'P'),adding_marks=lambda:False)
        def accept(_app,**kwargs):
            kwargs['on_accept']({'appendix_sha256':'DEV','indices':[]})
        with (patch('production_app.amendment.original_shipment',return_value=(None,None,original,historical,{})),
              patch('production_app.shipment_dialog.open_shipment',side_effect=accept)):
            App.add_marks(application)
        draft,=application.facts['marks']
        self.assertFalse(draft.get('uzk_length_mm') or draft.get('uzk_thickness_mm'))
        self.assertEqual(historical['marks'][0]['uzk_length_mm'],'155')


class ReportGeometry(unittest.TestCase):
    def test_summary_form_prints_all_groups_without_repeating_mark_or_total(self):
        self.check_summary_form(owner=False)

    def test_owner_summary_renders_before_confirmation_without_supplied_verdict(self):
        self.check_summary_form(owner=True)

    def check_summary_form(self, *, owner):
        # Native Office is exercised by the Windows release scenario. Here only
        # the COM boundary is stubbed so the real renderer chooses its template route.
        with TemporaryDirectory() as folder:
            root=Path(folder)
            templates={}
            docs={}
            for role in ('vik','uzk','quality'):
                source=root/(role+'.docx')
                doc=Document();doc.add_paragraph('{{NUMBER}}')
                if role=='uzk':doc.add_paragraph('{{UZK_TABLE}}')
                doc.save(source);templates[role]={'path':str(source)}
                docs[role]=Mock(Content=SimpleNamespace(Text='{{NUMBER}}'+ ('\r{{UZK_TABLE}}' if role=='uzk' else '')))
            mark={'mark':'M','uzk_quantity':2,'uzk_result':'DEV'}
            facts={'number':'2/2026','climates':[],'marks':[mark],
                   'uzk_design':[historical_part('75','8',origin='MANUAL'),historical_part('155','6')]}
            rows=(('M','DEV',2),)
            if owner:
                mark.pop('uzk_result')
                facts['owner_confirmation_required']=True
                facts['marks'].append({'mark':'Z','uzk_quantity':0})
                rows+=(('Z','DEV',1),)
            vom=SimpleNamespace(rows=rows)
            output=root/'out';output.mkdir()
            def open_doc(_app,path,**_kwargs):
                return docs[Path(path).stem.split(' - ')[1]]
            def finalize_text(doc,_values):
                doc.Content.Text='DEV'
            with (patch.object(office,'document_values',return_value={}),
                  patch.object(office,'excel_session'),patch.object(office,'word_session'),
                  patch.object(office,'open_word',side_effect=open_doc),
                  patch.object(office,'fill_word_controls',return_value=set()),
                  patch.object(office,'word_table') as table,
                  patch.object(office,'replace_word',side_effect=finalize_text),
                  patch.object(office,'word_preview_mark')):
                outputs=office.render({'templates':templates},None,None,None,vom,facts,output,training=True)
            _,token,headers,rows=table.call_args.args
            self.assertEqual(token,'{{UZK_TABLE}}')
            self.assertEqual(headers[-1],'Расчётная геометрия')
            self.assertEqual(len(rows),2 if owner else 1)
            self.assertEqual(rows[0][:3],('M',2,2))
            self.assertIn('L=75 мм, S=8 мм',rows[0][-1])
            self.assertIn('L=155 мм, S=6 мм',rows[0][-1])
            self.assertIn('дополнительный',rows[0][-1])
            self.assertIn('по реестру',rows[0][-1])
            if owner:
                self.assertEqual(rows[0][3:5],('Ожидает подтверждения владельца',''))
                self.assertEqual(rows[1][:5],('Z',1,0,'НЕ ТРЕБУЕТСЯ','План УЗК по текущей отгрузке: 0 шт.'))
                self.assertTrue(next(out for out in outputs if out['role']=='uzk')['owner_uzk_summary'])
            else:
                self.assertEqual(rows[0][3],'DEV')

    def test_other_design_routes_keep_all_weld_segments_under_one_mark(self):
        parts=[historical_part('75','8',origin='MANUAL'),historical_part('155','6',quantity=2)]
        for row in parts:row['quantity']=3
        self.assertEqual(_normalize_design(parts)['M'],[('75','8','5','70'),('155','6','5','70')])
        details={'mark':'M','welders':'DEV','date':'2026-11-26','defects':'DEV',
                 'conclusion':'DEV','basis':'DEV','frequency_mhz':'5','angle_deg':'70'}
        row,=native_rows({'marks':[{'mark':'M','uzk_quantity':3}],
                          'uzk_design':parts,'uzk_details':[details]})
        self.assertEqual(row[0],'M')
        total,manual,registered=row[1].splitlines()
        self.assertEqual(total,'Швеллер 16П; 3 шт.')
        self.assertIn('L=75х8 мм',manual)
        self.assertIn('дополнительный контроль, 1 шт.',manual)
        self.assertIn('L=155х6 мм',registered)
        self.assertIn('контроль по реестру, 2 шт.',registered)

    def test_exact_first_form_keeps_distinct_groups_when_one_row_is_full(self):
        parts=[]
        for number in range(6):
            parts.extend((historical_part(str(120+number),'5'),historical_part(str(240+number),'8.4')))
        for row in parts:row['quantity']=6
        with TemporaryDirectory() as folder:
            source=Path(folder)/'first.docx';target=Path(folder)/'issued.docx'
            uzk_first(source)
            create_uzk_copy(source,target,number='2/2026',uzk_date='2026-11-26',design_rows=parts)
            doc=Document(target)
        body=doc.tables[0].rows[5:]
        self.assertEqual([row.cells[0].text for row in body],['M','M'])
        self.assertTrue(all(row.cells[1]._tc is row.cells[2]._tc for row in body))
        groups=[line for row in body for line in row.cells[1].text.splitlines()]
        self.assertEqual(len(groups),12)
        self.assertEqual([line.split(' мм')[0] for line in groups],
                         [text for n in range(6) for text in (f'L={120+n}х5',f'L={240+n}х8,4')])
        self.assertTrue(all('контроль по реестру, 1 шт.' in line for line in groups))
        legacy=[{k:v for k,v in row.items() if k not in ('control_source','control_quantity')}
                for row in parts]
        with self.assertRaises(Stop):_normalize_design(legacy)
