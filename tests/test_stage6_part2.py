"""Variant isolation, source numerics, fixed-capacity and batch-publication boundaries."""
import copy
from decimal import Decimal
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch
from akz_stage6 import core as c,part2 as p

class Part2Boundaries(unittest.TestCase):
    def row(self,mark='A',rn=40):
        return dict(mark=mark,row=rn,name='Балка',quantity=2,mass_unit=3.25,mass_total=6.5,area_unit=1.25,area_total=2.5)
    def match(self,n,rows):return p.matching({'rows':rows},{'A':{'assignment':n,'row':2}},n)
    def test_three_exact_pairs_and_layer_maps(self):
        for n,cols in [(1,['E']),(2,['E','I']),(3,['E','H','J'])]:
            t=p.template(n)
            self.assertEqual([r['column'] for r in t['layers']],cols)
            self.assertEqual(Path(t['act']).name,f'АКЗ-{n}.xlsx')
            self.assertEqual(Path(t['application']).name,f'Приложение к АКЗ-{n}.xlsx')
    def test_cumulative_180_has_no_invented_80(self):
        second=p.template(2)['layers'][1]
        self.assertEqual(second['required_raw'],'∑1+2=180')
        self.assertIsNone(second['required_individual_microns'])
        self.assertEqual(second['requirement_status'],'NOT_PROVEN')
    def test_missing_exact_pair_never_falls_back(self):
        with tempfile.TemporaryDirectory() as d,patch.object(c,'PROJECT',Path(d)):
            (Path(d)/'Паспорт').mkdir()
            (Path(d)/'Паспорт/АКЗ-1 (копия).xlsx').write_bytes(b'NOT AN EXACT PAIR')
            with self.assertRaisesRegex(c.Stage6Error,'TEMPLATE_SHA_CHANGED_STOP'):p.template(1)
    def test_changed_template_pin_stop(self):
        with patch.object(p,'PINS',{**p.PINS,1:('0'*64,p.PINS[1][1])}):
            with self.assertRaisesRegex(c.Stage6Error,'TEMPLATE_SHA_CHANGED_STOP'):p.template(1)
    def test_numeric_grouping_without_cross_mixing_or_blank_zero(self):
        rows=[self.row('ZERO',1),self.row('A',2),self.row('B',3),self.row('C',4),self.row('BLANK',5)]
        assignments={m:{'assignment':n,'row':i+2} for i,(m,n) in enumerate([('ZERO',0),('A',1),('B',2),('C',3),('BLANK',None)])}
        for n,mark in [(1,'A'),(2,'B'),(3,'C')]:
            selected,excluded,_=p.matching({'rows':rows},assignments,n)
            self.assertEqual([r['mark'] for r in selected],[mark]);self.assertEqual(len(excluded),4)
        with self.assertRaisesRegex(c.Stage6Error,'NO_VARIANT_ROWS'):p.matching({'rows':[self.row('BLANK')]},assignments,1)
    def test_repeated_rows_order_and_exact_source_totals(self):
        for n in (1,2,3):
            selected,_,totals=self.match(n,[self.row(rn=53),self.row(rn=40),self.row(rn=54)])
            self.assertEqual([r['row'] for r in selected],[53,40,54])
            self.assertEqual(totals,dict(quantity=Decimal(6),mass_total=Decimal('19.5'),area_total=Decimal('7.5')))
            self.assertEqual(selected[0]['mass_unit'],Decimal('3.25'));self.assertEqual(selected[0]['area_unit'],Decimal('1.25'))
    def test_all_capacity_boundaries(self):
        for n,cap in [(1,61),(2,47),(3,94)]:
            self.assertEqual(len(self.match(n,[self.row(rn=i) for i in range(cap)])[0]),cap)
            with self.assertRaisesRegex(c.Stage6Error,'CAPACITY_STOP'):
                self.match(n,[self.row(rn=i) for i in range(cap+1)])
    def test_missing_or_disputed_area_blocks_only_its_group(self):
        for n in (1,2,3):
            for field,value,code in [('area_unit',None,'NUMBER_MISSING'),('area_total',None,'NUMBER_MISSING'),('area_total',3,'AREA_CONFLICT_STOP')]:
                with self.assertRaisesRegex(c.Stage6Error,code):self.match(n,[{**self.row(),field:value}])
        rows=[self.row(),{**self.row('OTHER'), 'area_total':None}]
        self.assertEqual(len(self.match(1,rows)[0]),1)
    def test_invalid_numeric_values_not_coerced(self):
        for value in (True,'1',float('nan'),0,-1):
            with self.assertRaises(c.Stage6Error):self.match(1,[{**self.row(),'area_unit':value}])
    def test_official_stop_before_input_or_files(self):
        with self.assertRaisesRegex(c.Stage6Error,'OFFICIAL_MODE_STOP'):p.generate(None,'x',mode='OFFICIAL')
    def test_exact_safe_output_names(self):
        self.assertEqual(p.filenames('9/2026',2),('9_2026-2 АКЗ TEST.xlsx','9_2026-2 Приложение АКЗ TEST.xlsx','PAIR.json'))
        for passport in ('../9/2026','9/2026/1','9:2026'):
            with self.assertRaises(c.Stage6Error):p.filenames(passport,1)
    def test_batch_tampering_stop(self):
        data={'groups':{},'sources':{}};batch=p.BatchPreview(data,c.digest(data));data['groups']['1']={}
        with self.assertRaisesRegex(c.Stage6Error,'BATCH_PREVIEW_CHANGED_STOP'):p.group_preview(batch,1)
    def test_input_changed_stop(self):
        with tempfile.TemporaryDirectory() as d:
            path=Path(d)/'input';path.write_bytes(b'A')
            data={'sources':{str(path):c.sha256(path)},'groups':{}}
            batch=p.BatchPreview(data,c.digest(data));path.write_bytes(b'B')
            with self.assertRaisesRegex(c.Stage6Error,'INPUT_CHANGED_AFTER_PREVIEW_STOP'):p.group_preview(batch,1)
    def test_selection_preflight_prevents_publishing_invalid_group(self):
        data={'sources':{},'groups':{'1':{'status':'STOP','reason':'missing area'}}};batch=p.BatchPreview(data,c.digest(data))
        with patch.object(p,'generate') as generate:
            with self.assertRaisesRegex(c.Stage6Error,'GROUP_NOT_READY'):p.generate_groups(batch,[1],'test')
            generate.assert_not_called()
    def test_partial_batch_never_reports_full_success(self):
        with patch.object(p,'group_preview',return_value=object()),patch.object(p,'generate',side_effect=[Path('pair1'),c.Stage6Error('FILE_LOCKED')]):
            result=p.generate_groups(None,[1,2],'test')
            self.assertEqual(result['status'],'PARTIAL_STOP')
            self.assertEqual(result['groups']['1']['status'],'COMPLETE_TEST_PAIR')
            self.assertEqual(result['groups']['2']['status'],'STOP')
    def test_group_failure_isolated_and_single_daily_climate(self):
        rows=[self.row('A'),{**self.row('B'),'area_total':None},self.row('C')]
        assignments={mark:{'assignment':n,'row':n+1} for n,mark in enumerate(('A','B','C'),1)}
        with patch.object(c,'current_passport',return_value={'date':'2026-10-05','number':'9/2026'}),\
             patch.object(c,'catalog',return_value=assignments),\
             patch.object(c,'visible_vom',return_value={'rows':rows,'synthetic_source':False}),\
             patch.object(c,'sha256',return_value='unchanged'),patch.object(p,'template',return_value={}):
            batch=p.preview_groups('Гараж Казахстан','9/2026','vom.xls','catalog.xlsx')
        groups=batch.payload['groups']
        self.assertEqual([groups[str(n)]['status'] for n in range(4)],['EMPTY','READY','STOP','READY'])
        self.assertEqual(groups['1']['payload']['climate'],groups['3']['payload']['climate'])
        self.assertEqual(list(batch.payload['climate_days']),['2026-10-05'])
        self.assertEqual(groups['1']['payload']['climate']['confirmation'],'MISSING')
    def test_registry_change_during_read_stop(self):
        state={'hash':'BEFORE'}
        def current(*args):
            state['hash']='AFTER'
            return {'date':'2026-10-05','number':'9/2026'}
        with patch.object(c,'sha256',side_effect=lambda path:state['hash']),\
             patch.object(c,'current_passport',side_effect=current),\
             patch.object(c,'catalog',return_value={}),\
             patch.object(c,'visible_vom',return_value={'rows':[],'synthetic_source':False}):
            with self.assertRaisesRegex(c.Stage6Error,'INPUT_CHANGED_DURING_PREVIEW'):
                p.preview_groups('Гараж Казахстан','9/2026','vom.xls','catalog.xlsx')

if __name__=='__main__':unittest.main()
