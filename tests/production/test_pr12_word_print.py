"""Word UZK isolation and Excel print-title unit contracts (no Office needed)."""
import unittest
from types import SimpleNamespace

from production_app.office import word_table, set_appendix_print_titles, verify_appendix_print_layout, fit_appendix_body
from production_app.storage import Stop

TOKEN='{{UZK_DETAILS_TABLE}}'
HEADERS=tuple('HEAD_'+str(n) for n in range(1,9))
ROWS=(tuple('VALUE_'+str(n) for n in range(1,9)),)


class FakeAnchor:
    def __init__(self,doc):
        self.doc=doc
        self.Find=SimpleNamespace(Execute=lambda **kw: kw.get('FindText')==TOKEN
                                 and not self.doc.removed)
        self.Tables=SimpleNamespace(Count=int(doc.marker_in_table))
        self.Paragraphs=FakeParagraphs(doc)

    @property
    def Text(self):
        return TOKEN if not self.doc.removed else ''

    @Text.setter
    def Text(self,value):
        if value!='':
            raise AssertionError('Unexpected table token replacement')
        self.doc.removed=True


class FakeParagraphs:
    def __init__(self,doc):
        self.doc=doc
        self.Count=1

    def __call__(self,index):
        assert index==1
        paragraph=SimpleNamespace(
            Text=(TOKEN+' extra' if self.doc.extra_text else TOKEN+'\r'),
            InsertParagraphBefore=lambda: setattr(self.doc,'separator',True),
        )
        return SimpleNamespace(Range=paragraph)


class FakeContent:
    def __init__(self,doc):
        self.doc=doc
    @property
    def Text(self):
        return TOKEN if not self.doc.removed else ''
    @property
    def Duplicate(self):
        return FakeAnchor(self.doc)


class FakeTable:
    def __init__(self,rows,columns):
        self.rows=rows
        self.columns=columns
        self.Borders=SimpleNamespace(Enable=None)
        self.values={}
        self.fit_mode=None
    def AutoFitBehavior(self,mode):
        self.fit_mode=mode
    def Rows(self,index):
        assert 1<=index<=self.rows
        return SimpleNamespace(Cells=SimpleNamespace(Count=self.columns),
                               HeadingFormat=None,Range=SimpleNamespace(
                                   Font=SimpleNamespace(Bold=False)))
    def Cell(self,row,column):
        assert 1<=row<=self.rows
        assert 1<=column<=self.columns
        key=(row,column)
        if key not in self.values:
            self.values[key]=SimpleNamespace(Text='')
        return SimpleNamespace(Range=self.values[key])


class FakeTables:
    def __init__(self,doc):
        self.doc=doc
        self.created=None
    @property
    def Count(self):
        return 2 if self.created is None or self.doc.force_join else 3
    def Add(self,anchor,rows,columns):
        self.created=FakeTable(rows,columns if self.doc.separator else 2)
        return self.created


class FakeWord:
    def __init__(self,*,marker_in_table=False,extra_text=False,force_join=False):
        self.removed=False
        self.separator=False
        self.marker_in_table=marker_in_table
        self.extra_text=extra_text
        self.force_join=force_join
        self.Content=FakeContent(self)
        self.Tables=FakeTables(self)


class FakeExcelRange:
    def __init__(self, *, merged=False):
        self.MergeCells=merged
        self.WrapText=False
        self.ShrinkToFit=True
        self.VerticalAlignment=None


class FakeExcelRows:
    def __init__(self,sheet,index):
        self.sheet=sheet
        self.index=str(index)
        self.start=int(self.index.split(':')[0])
        self.end=int(self.index.split(':')[-1])

    @property
    def RowHeight(self):
        return self.sheet.heights[self.start]
    @RowHeight.setter
    def RowHeight(self,value):
        for row in range(self.start,self.end+1):
            self.sheet.heights[row]=value

    @property
    def Hidden(self):
        return self.sheet.hidden.get(self.start,False)

    def AutoFit(self):
        self.sheet.called_autofit=True
        for row in range(self.start,self.end+1):
            text=' '.join(str(self.sheet.values.get((row,col),''))
                          for col in (1,2))
            # Synthetic Excel estimate for long wrapped input.
            self.sheet.heights[row]=18+(len(text)//35)*14


class FakeExcelSheet:
    def __init__(self,start):
        self.start=start
        self.values={
            (1,1):'ПРИЛОЖЕНИЕ — длинный номер и дата',
            (2,1):'ОБЪЕКТ — очень длинное наименование объекта',
            (max(3,start-2),1):'Отправочная марка',
            (start,1):'МАРКА-001',
        }
        self.PageSetup=SimpleNamespace(PrintTitleRows='$1:$4',PrintArea='$A$1:$G$110',
                                       FitToPagesWide=None,
                                       FitToPagesTall=None,Zoom=None)
        self.heights={i:15.0 for i in range(1,320)}
        self.hidden={}
        self.ranges={}
        self.called_autofit=False
    def Cells(self,row,col):
        return SimpleNamespace(Value2=self.values.get((row,col)))
    def Rows(self,index):
        return FakeExcelRows(self,index)
    def Range(self,address):
        if address not in self.ranges:
            self.ranges[address]=FakeExcelRange()
        return self.ranges[address]


class OfficeContracts(unittest.TestCase):
    def test_uzk_has_new_independent_eight_column_table(self):
        doc=FakeWord()
        table=word_table(doc,TOKEN,HEADERS,ROWS)
        self.assertTrue(doc.separator)
        self.assertEqual(doc.Tables.Count,3)
        self.assertEqual(table.Rows(1).Cells.Count,8)
        self.assertEqual(table.rows,2)
        self.assertEqual(table.fit_mode,2)
        self.assertEqual([table.Cell(1,n).Range.Text for n in range(1,9)],
                         list(HEADERS))
        self.assertEqual([table.Cell(2,n).Range.Text for n in range(1,9)],
                         list(ROWS[0]))

    def test_uzk_marker_inside_requisites_is_refused(self):
        doc=FakeWord(marker_in_table=True)
        with self.assertRaisesRegex(Stop,'внутри другой таблицы'):
            word_table(doc,TOKEN,HEADERS,ROWS)
        self.assertFalse(doc.removed)

    def test_uzk_marker_must_be_standalone_paragraph(self):
        doc=FakeWord(extra_text=True)
        with self.assertRaisesRegex(Stop,'единственным текстом'):
            word_table(doc,TOKEN,HEADERS,ROWS)

    def test_word_join_is_detected_before_cell_population(self):
        doc=FakeWord(force_join=True)
        with self.assertRaisesRegex(Stop,'присоединил'):
            word_table(doc,TOKEN,HEADERS,ROWS)

    def test_wrong_row_width_refused_in_copy(self):
        doc=FakeWord()
        with self.assertRaisesRegex(Stop,'неверное число колонок'):
            word_table(doc,TOKEN,HEADERS,ROWS+(('bad',),))

    def test_appendix_vzmk_two_tier_headers_exclude_title_and_object(self):
        sheet=FakeExcelSheet(start=5)
        self.assertEqual(set_appendix_print_titles(sheet,5),(3,4))
        self.assertEqual(sheet.PageSetup.PrintTitleRows,'$3:$4')
        self.assertEqual(sheet.PageSetup.FitToPagesWide,1)
        self.assertEqual(sheet.PageSetup.FitToPagesTall,False)
        self.assertEqual(sheet.values[(1,1)],
                         'ПРИЛОЖЕНИЕ — длинный номер и дата')
        self.assertIn('очень длинное',sheet.values[(2,1)])

    def test_tall_repeated_headers_avoid_page_three_clipping(self):
        sheet=FakeExcelSheet(start=5)
        self.assertEqual(sheet.Rows(3).RowHeight,15)
        set_appendix_print_titles(sheet,5)
        self.assertGreaterEqual(sheet.Rows(3).RowHeight,28)
        self.assertGreaterEqual(sheet.Rows(4).RowHeight,28)
        self.assertTrue(sheet.Range('A3:G4').WrapText)
        self.assertEqual(sheet.PageSetup.PrintTitleRows,'$3:$4')

    def test_long_mark_names_trigger_autofit_not_fixed_18pt(self):
        sheet=FakeExcelSheet(start=5)
        sheet.values[(6,2)]='Тестовое наименование марки с большим количеством слов '*3
        self.assertEqual(fit_appendix_body(sheet,5,8),4)
        self.assertTrue(sheet.called_autofit)
        self.assertGreater(sheet.Rows(6).RowHeight,18)
        self.assertTrue(sheet.Range('A5:B8').WrapText)
        self.assertFalse(sheet.Range('A5:B8').ShrinkToFit)

    def test_merged_or_overlong_data_row_stops_before_print(self):
        sheet=FakeExcelSheet(start=5)
        sheet.Range('A5:B8').MergeCells=True
        with self.assertRaisesRegex(Stop,'объединённые'):
            fit_appendix_body(sheet,5,8)
        sheet=FakeExcelSheet(start=5)
        sheet.values[(6,2)]='very long synthetic label '*70
        with self.assertRaisesRegex(Stop,'Слишком длинное'):
            fit_appendix_body(sheet,5,8)

    def test_hidden_source_row_stops_instead_of_omitting_marks(self):
        sheet=FakeExcelSheet(start=5)
        sheet.hidden[6]=True
        with self.assertRaisesRegex(Stop,'скрыта'):
            fit_appendix_body(sheet,5,8)

    def test_generic_one_tier_header_excludes_object(self):
        sheet=FakeExcelSheet(start=4)
        self.assertEqual(set_appendix_print_titles(sheet,4),(3,3))
        self.assertEqual(sheet.PageSetup.PrintTitleRows,'$3:$3')

    def test_invalid_header_contract_stop(self):
        sheet=FakeExcelSheet(start=5)
        sheet.values[(3,1)]=''
        with self.assertRaisesRegex(Stop,'Нет заголовка'):
            set_appendix_print_titles(sheet,5)

    def test_reopening_with_old_repeated_title_is_refused(self):
        sheet=FakeExcelSheet(start=5)
        sheet.PageSetup.PrintTitleRows='$1:$4'
        with self.assertRaisesRegex(Stop,'длинные строки'):
            verify_appendix_print_layout(sheet)

    def test_reopened_compact_print_titles_preserved(self):
        sheet=FakeExcelSheet(start=5)
        sheet.PageSetup.PrintTitleRows="'Приложение №1'!$3:$4"
        sheet.heights[3]=28
        sheet.heights[4]=28
        self.assertEqual(verify_appendix_print_layout(sheet,'3:4'),'3:4')

    def test_page_three_header_too_short_after_reopen_is_refused(self):
        sheet=FakeExcelSheet(start=5)
        sheet.PageSetup.PrintTitleRows='$3:$4'
        # Correct PrintTitleRows alone never certifies print legibility.
        with self.assertRaisesRegex(Stop,'слишком низкий'):
            verify_appendix_print_layout(sheet,'3:4')

    def test_missing_print_area_refused(self):
        sheet=FakeExcelSheet(start=5)
        sheet.PageSetup.PrintTitleRows='$3:$4'
        sheet.PageSetup.PrintArea=''
        with self.assertRaisesRegex(Stop,'область печати'):
            verify_appendix_print_layout(sheet)

    def test_unexpected_print_title_reset_is_rejected(self):
        sheet=FakeExcelSheet(start=5)
        class BrokenPageSetup:
            FitToPagesWide=None
            FitToPagesTall=None
            Zoom=None
            PrintArea='$A$1:$G$110'
            @property
            def PrintTitleRows(self):return '$1:$4'
            @PrintTitleRows.setter
            def PrintTitleRows(self,value):pass
        sheet.PageSetup=BrokenPageSetup()
        with self.assertRaisesRegex(Stop,'не сохранил'):
            set_appendix_print_titles(sheet,5)


if __name__=='__main__':
    unittest.main(verbosity=2)
