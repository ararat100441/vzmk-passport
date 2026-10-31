import unittest
from decimal import Decimal
from types import SimpleNamespace
from unittest.mock import patch
from production_app.office import document_values
from production_app.storage import Stop


class PrintedShipmentValues(unittest.TestCase):
    def setUp(self):
        self.facts={'number':'11/2026','responsible':'TRAINING','evidence':[],
                    'start_date':'2026-10-01','vik_date':'2026-10-17',
                    'uzk_date':'2026-10-19','akz_date':'2026-10-23','climates':[]}
        self.values={'MASS':'15822.14','START_DATE':'01.10.2026',
                     'VIK_DATE':'17.10.2026','UZK_DATE':'19.10.2026',
                     'AKZ_DATE':'23.10.2026'}

    def values_for(self,vom):
        with patch('production_app.office.fields',return_value=self.values.copy()), \
                patch('production_app.office.values_for_documents',return_value={}):
            return document_values(None,1,vom,self.facts)

    def test_printed_mass_matches_verified_registry_not_binary_sum_tail(self):
        raw=Decimal('15822.13999999999951')
        vom=SimpleNamespace(mass=raw,connections='Балка',rows=(('M','Балка',1,1,1,1,1),))
        result=self.values_for(vom)
        self.assertEqual(result['MASS'],'15822.14')
        self.assertEqual(vom.mass,raw)
        self.assertEqual(result['CONNECTIONS'],'Балка')
        self.assertEqual(result['VIK_PROCESS_NUMBER'],'11/2026-1')

    def test_registry_precision_is_preserved_without_new_rounding_rule(self):
        self.values['MASS']='12.345'
        self.assertEqual(self.values_for(SimpleNamespace(mass=Decimal('12.345'),connections='A',rows=(('M','A',1,1,1,1,1),)))['MASS'],'12.345')

    def test_inconsistent_control_date_remains_a_hard_refusal(self):
        self.facts['vik_date']='2026-10-16'
        with self.assertRaises(Stop):
            self.values_for(SimpleNamespace(mass=Decimal('15822.14'),connections='A',rows=(('M','A',1,1,1,1,1),)))
