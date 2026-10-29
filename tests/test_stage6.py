"""Meaningful Stage 6 boundary tests; mutations use in-memory isolated fixtures."""
import copy
from dataclasses import replace
from decimal import Decimal
import unittest
from unittest.mock import patch
from akz_stage6 import core as c

class Stage6Boundaries(unittest.TestCase):
    def row(self):
        return {'row':4,'mark':'ТЕСТ-1','name':'TEST','quantity':2,
                'mass_unit':3.25,'mass_total':6.5,'area_unit':1.25,'area_total':2.5}
    def match(self,rows,assignment=0):
        return c.matching({'rows':rows},{'ТЕСТ-1':{'assignment':assignment,'row':2}})
    def test_numeric_zero_and_blank_separately(self):
        self.assertEqual(len(self.match([self.row()],0)[0]),1)
        with self.assertRaisesRegex(c.Stage6Error,'NO_AKZ_ZERO_ROWS'):self.match([self.row()],None)
    def test_variants_1_2_3_excluded(self):
        for n in (1,2,3):
            with self.subTest(variant=n),self.assertRaisesRegex(c.Stage6Error,'NO_AKZ_ZERO_ROWS'):
                self.match([self.row()],n)
    def test_duplicates_and_vom_order(self):
        rows=[self.row(),{**self.row(),'row':8},{**self.row(),'row':12}]
        selected,_,totals=self.match(rows)
        self.assertEqual([r['row'] for r in selected],[4,8,12])
        self.assertEqual(totals,{'quantity':Decimal(6),'mass_total':Decimal('19.5'),'area_total':Decimal('7.5')})
    def test_missing_area_stop(self):
        for f in ('area_unit','area_total'):
            with self.subTest(field=f),self.assertRaisesRegex(c.Stage6Error,'NUMBER_MISSING'):
                self.match([{**self.row(),f:None}])
    def test_conflicting_area_stop(self):
        with self.assertRaisesRegex(c.Stage6Error,'AREA_CONFLICT_STOP'):
            self.match([{**self.row(),'area_total':25}])
    def test_rounding_not_invented(self):
        with self.assertRaisesRegex(c.Stage6Error,'AREA_CONFLICT_STOP'):
            self.match([{**self.row(),'area_total':2.501}])
    def test_missing_mass_stop(self):
        with self.assertRaisesRegex(c.Stage6Error,'NUMBER_MISSING'):
            self.match([{**self.row(),'mass_unit':None}])
    def test_conflicting_mass_stop(self):
        with self.assertRaisesRegex(c.Stage6Error,'MASS_CONFLICT_STOP'):
            self.match([{**self.row(),'mass_total':65}])
    def test_exact_mark_no_fuzzy_match(self):
        with self.assertRaisesRegex(c.Stage6Error,'NO_AKZ_ZERO_ROWS'):
            self.match([{**self.row(),'mark':'ТЕСТ-1 '}])
    def test_other_unassigned_area_does_not_block(self):
        rows=[self.row(),{**self.row(),'mark':'OTHER','area_total':None}]
        selected,excluded,_=self.match(rows)
        self.assertEqual(len(selected),1);self.assertEqual(excluded[0]['reason'],'NO_ASSIGNMENT')
    def test_dynamic_application_counts_no_truncation(self):
        for count in (1,7,17,18,104,200):
            with self.subTest(count=count):
                rows=[{**self.row(),'row':i+4} for i in range(count)]
                selected,_,totals=self.match(rows)
                self.assertEqual([r['row'] for r in selected],[r['row'] for r in rows])
                self.assertEqual(totals,dict(quantity=Decimal(2*count),
                    mass_total=Decimal('6.5')*count,area_total=Decimal('2.5')*count))
    def test_numeric_rejects_bool_nan_negative_and_text(self):
        for v in (True,float('nan'),-1,'0'):
            with self.subTest(value=v),self.assertRaises(c.Stage6Error):c.number(v,'TEST')
    def test_climate_one_per_day_and_boundary(self):
        record=c.ClimateDay('2026-10-05',Decimal(20),Decimal(50),Decimal(18),Decimal(15),'TEST')
        days={};c.register_climate(days,record);c.register_climate(days,record)
        self.assertEqual(len(days),1)
        with self.assertRaisesRegex(c.Stage6Error,'CLIMATE_DAY_CONFLICT'):
            c.register_climate(days,replace(record,air_c=Decimal(21)))
        with self.assertRaisesRegex(c.Stage6Error,'CLIMATE_DEW_MARGIN_STOP'):
            replace(record,surface_c=Decimal('17.999')).validate()
        c.ClimateDay('2026-10-05').validate()
    def test_climate_no_silent_formula_or_partial_readings(self):
        with self.assertRaisesRegex(c.Stage6Error,'CLIMATE_INCOMPLETE'):
            c.ClimateDay('2026-10-05',air_c=Decimal(20),confirmation='TEST').validate()
    def test_official_always_stop_before_read(self):
        with self.assertRaisesRegex(c.Stage6Error,'OFFICIAL_MODE_STOP'):c.generate(None,'TEST',mode='OFFICIAL')
    def test_preview_payload_tampering_stop(self):
        data={'mode':'TEST_ONLY','sources':{},'official':'STOP'}
        p=c.Preview(data,c.digest(data));data['official']='YES'
        with self.assertRaisesRegex(c.Stage6Error,'PREVIEW_CHANGED_STOP'):c.verify_preview(p)

if __name__=='__main__':unittest.main()
