"""Pure structural contracts for Windows-only historical АКЗ adapters.

Fixtures are synthetic and contain no real client or inspection data.
Microsoft Excel COM round-trip remains a separate Windows acceptance gate.
"""
import unittest
from types import SimpleNamespace
from tools.prepare_akz_three_layer_template import (
    validate_three_layer_source, verify_converted_three_layer_sheet,
)
from tools.prepare_akz_appendix_template import validate_converted_appendix


class Cell:
    def __init__(self, value, merged):
        self.Value2=value
        self.MergeArea=SimpleNamespace(Address=merged)


class Sheet:
    def __init__(self, values=None, merges=None, area='$A$1:$M$50'):
        self.values=dict(values or {})
        self.merges=dict(merges or {})
        self.PageSetup=SimpleNamespace(PrintArea=area)
    def Range(self, address):
        return Cell(self.values.get(address),
                    self.merges.get(address,'$'+address.replace('1','$1')))


def source():
    values={
        'B3':'СЕРТИФИКАТ КАЧЕСТВА АНТИКОРРОЗИОННОЙ ОБРАБОТКИ',
        'B17':'Слой №','B18':'Наименование ЛКМ',
        'B19':'Серийный № краски','B20':'Серийный № отвердителя',
        'B22':'Дата','C25':'Требуемая толщина, мкм',
        'C28':'Средняя толщина сухого покрытия, мкм',
        'B31':'Температура поверхности','B32':'Температура воздуха',
        'B33':'Уровень влажности','B34':'Точка росы',
    }
    merges={}
    for col,n,end in (('E',1,'G'),('H',2,'I'),('J',3,'K')):
        values[col+'17']=n
        values[col+'18']='Synthetic paint '+str(n)
        values[col+'19']='Synthetic batch'
        values[col+'20']='Synthetic hardener batch'
        values[col+'22']=46000+n
        values[col+'27']=110*n
        values[col+'28']=100*n
        merges[col+'18']='$'+col+'$18:$'+end+'$18'
    return Sheet(values,merges)


def converted():
    values={'E7':'{{OBJECT}}'}
    for col,n,_ in (('E',1,'G'),('H',2,'I'),('J',3,'K')):
        values[col+'17']=n
        prefix='AKZ_LAYER_'+str(n)+'_'
        for row,name in (
                (18,'MATERIAL'),(19,'PAINT_BATCH'),(20,'HARDENER_BATCH'),
                (21,'COLOR'),(22,'APPLICATION_DATE'),(23,'TIME_START'),
                (24,'TIME_END'),(27,'WET_AVG')):
            values[col+str(row)]='{{'+prefix+name+'}}'
        for row,name in ((25,'RANGE'),(26,'TOTAL_RANGE'),(28,'DRY_AVG')):
            values[col+str(row)]='{{LAYER_'+str(n)+'_'+name+'}}'
        for row,name in ((31,'SURFACE'),(32,'AIR'),(33,'HUMIDITY'),(34,'DEW')):
            values[col+str(row)]='{{'+name+'}}'
    return Sheet(values)


def appendix():
    values={'A1':'Приложение № {{NUMBER}} от {{AKZ_DATE}}',
            'A2':'Объект: {{OBJECT}}','A3':'Отправочная марка',
            'A5':'{{AKZ_MARKS}}','A6':'Итого'}
    return Sheet(values,{'A2':'$A$2:$G$2'},'$A$1:$G$6')


class AdapterContracts(unittest.TestCase):
    def test_three_layers_accepted_by_strict_signature(self):
        self.assertTrue(validate_three_layer_source(source()))

    def test_single_layer_cannot_pass_due_to_title_or_label(self):
        single=source()
        single.values['H17']=None
        single.values['J17']=None
        with self.assertRaisesRegex(ValueError,'слой №2'):
            validate_three_layer_source(single)

    def test_missing_independent_coat_value_refused(self):
        for address in ('H18','J20','H22','J28'):
            sample=source()
            sample.values[address]=None
            with self.subTest(address=address),self.assertRaises(ValueError):
                validate_three_layer_source(sample)

    def test_different_merge_layout_not_silently_used(self):
        sample=source()
        sample.merges['H18']='$H$18:$J$18'
        with self.assertRaisesRegex(ValueError,'объедин'):
            validate_three_layer_source(sample)

    def test_candidate_has_all_fields_after_save(self):
        self.assertTrue(verify_converted_three_layer_sheet(converted()))
        sample=converted()
        sample.values['J28']='Historical mean'
        with self.assertRaisesRegex(ValueError,'толщины'):
            verify_converted_three_layer_sheet(sample)

    def test_appendix_has_object_and_dynamic_row(self):
        self.assertTrue(validate_converted_appendix(appendix()))

    def test_appendix_missing_object_or_header_or_print_refused(self):
        for address,value in (('A2',''),('A3','not a mark header'),
                              ('A5',''),('A6','old total'),('A1','other')):
            sample=appendix()
            sample.values[address]=value
            with self.subTest(address=address),self.assertRaises(ValueError):
                validate_converted_appendix(sample)
        sample=appendix()
        sample.PageSetup.PrintArea=''
        with self.assertRaisesRegex(ValueError,'печати'):
            validate_converted_appendix(sample)

    def test_appendix_object_must_be_merged(self):
        sample=appendix()
        sample.merges['A2']='$A$2:$D$2'
        with self.assertRaisesRegex(ValueError,'A2:G2'):
            validate_converted_appendix(sample)


if __name__=='__main__':
    unittest.main(verbosity=2)
