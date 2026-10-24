"""Create a separate, UNAPPROVED, three-layer Excel АКЗ form.

Run on Windows with locally installed Microsoft Excel. The historical issued
document is READ ONLY; all work is on a newly named private copy. Never
assume that old measurements, signatures, batches or dates are valid.

This conversion deliberately replaces old wet-thickness formulas with actual
confirmed wet-mean inputs; the previously embedded solids fraction is cleared.
Independent native Excel visual approval is mandatory before production use.
"""
import argparse
import hashlib
import json
import os
import shutil
from pathlib import Path
from uuid import uuid4


def sha(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


LAYERS = (('E',1),('H',2),('J',3))
STATIC = {
    'E4': '{{NUMBER}} от ', 'H4': '{{AKZ_DATE}}',
    'E5': '{{CONTRACT}}', 'E6': '{{CUSTOMER}}', 'E7': '{{OBJECT}}',
    'E13': '{{AKZ_DEGREASING}}', 'I13': '{{AKZ_DUST_REMOVAL}}',
    'B13': 'Rz = {{AKZ_SURFACE_ROUGHNESS}} мкм',
    'B9': 'Степень очистки поверхности по ISO 8501-1: {{AKZ_CLEANING_GRADE}}',
    'B14': 'Подготовка поверхности по ISO 8501-3: {{AKZ_PREPARATION_GRADE}}',
    'B35': 'Способ нанесения ЛКМ: {{AKZ_APPLICATION_METHOD}}',
    'L27': None,  # legacy, unconfirmed volume solids percentage
    'E39': '{{AKZ_HYGROMETER}}',
    'E40': '{{AKZ_THICKNESS_GAUGE}}',
    'E41': '{{AKZ_COMPARATOR}}',
    'E42': '{{AKZ_PROFILOMETER}}',
    'B44': '{{CIPHER}}', 'E44': '{{QUALITY_STANDARDS}}',
    'E46': '{{AKZ_QC_INSPECTOR}}', 'E47': '{{AKZ_PAINT_SHOP_HEAD}}',
    'F50': '{{AKZ_QC_INSPECTOR}}', 'I50': '{{AKZ_PAINT_SHOP_HEAD}}',
    # Old selection marks represented surface grade & application method.
    # Never carry those approvals from the old act into a new shipment.
    'G11': None, 'E15': None, 'G37': None,
}


def validate_three_layer_source(sheet):
    """Reject single-layer imitations of the audited VZMK 3-layer act."""
    def val(address):
        return sheet.Range(address).Value2
    def label(address):
        return ' '.join(str(val(address) or '').casefold().replace('ё','е').split())
    required={'B17':'слой','B18':'наименование лкм','B19':'серийный',
              'B20':'отвердител','B22':'дата','C25':'требуемая толщина',
              'C28':'средняя толщина','B31':'температура поверхности',
              'B32':'температура воздуха','B33':'влажности','B34':'точка росы'}
    for address,needle in required.items():
        if needle not in label(address):
            raise ValueError('Неизвестная трёхслойная структура: '+address)
    if 'сертификат качества' not in label('B3'):
        raise ValueError('Не найдена форма сертификата АКЗ.')
    for col,index in LAYERS:
        raw=val(col+'17')
        if type(raw) not in (int,float) or raw!=index:
            raise ValueError('Не найден самостоятельный слой №'+str(index)
                             +' в '+col+'17')
        for row in (18,19,20,22,27,28):
            v=val(col+str(row))
            if v in (None,''):
                raise ValueError('Нет данных слоя №'+str(index)+': '+col+str(row))
        for row in (22,27,28):
            v=val(col+str(row))
            if type(v) not in (int,float) or not 0<v<100000:
                raise ValueError('Неверное значение '+col+str(row))
        merge=str(sheet.Range(col+'18').MergeArea.Address).replace('
    source=Path(source).resolve(strict=True)
    dest=Path(destination).absolute()
    if source.suffix.lower()!='.xlsx' or dest.suffix.lower()!='.xlsx':
        raise ValueError('Требуется исходный и новый файл XLSX.')
    if dest.exists() or source==dest:
        raise FileExistsError('Новый шаблон должен иметь ещё не существующее имя.')
    if not dest.parent.is_dir():
        raise FileNotFoundError('Создайте отдельную папку для шаблона.')
    original=sha(source)
    scratch=dest.parent/('.vzmk-three-coat-'+uuid4().hex+'.xlsx')
    shutil.copyfile(source,scratch)
    app=None;book=None
    try:
        import pythoncom
        import win32com.client
        pythoncom.CoInitialize()
        try:
            app=win32com.client.DispatchEx('Excel.Application')
            app.Visible=False;app.DisplayAlerts=False;app.EnableEvents=False
            app.AskToUpdateLinks=False;app.AutomationSecurity=3
            book=app.Workbooks.Open(Filename=str(scratch),UpdateLinks=0,
                                    ReadOnly=False,AddToMru=False)
            if book.Worksheets.Count!=1 or book.ReadOnly:
                raise RuntimeError('Ожидался один редактируемый лист.')
            sheet=book.Worksheets(1)
            if sheet.ProtectContents:
                raise RuntimeError('Лист защищён.')
            validate_three_layer_source(sheet)

            # Only edit the fresh copy, leaving merges, dimensions, image,
            # print areas, page setup and all approved static labels in place.
            for cell,token in STATIC.items():
                sheet.Range(cell).Value2=token
            for column,n in LAYERS:
                pref='AKZ_LAYER_'+str(n)+'_'
                fields={
                    18:pref+'MATERIAL',
                    19:pref+'PAINT_BATCH',
                    20:pref+'HARDENER_BATCH',
                    21:pref+'COLOR',
                    22:pref+'APPLICATION_DATE',
                    23:pref+'TIME_START',
                    24:pref+'TIME_END',
                    25:'LAYER_'+str(n)+'_RANGE',
                    26:'LAYER_'+str(n)+'_TOTAL_RANGE',
                    27:pref+'WET_AVG',
                    28:'LAYER_'+str(n)+'_DRY_AVG',
                    31:'SURFACE',32:'AIR',33:'HUMIDITY',34:'DEW',
                }
                for row,field in fields.items():
                    cell=sheet.Range(column+str(row))
                    cell.Value2='{{'+field+'}}'
            # Keep technical markers legible for Windows print review.
            for cell in ('E46','E47','F50','I50'):
                sheet.Range(cell).WrapText=True
            sheet.Rows('46:47').RowHeight=45
            sheet.Rows('50:50').RowHeight=78
            verify_converted_three_layer_sheet(sheet)
            # Existing day formulas rely on native date serials at row 22;
            # those will be restored by production_app.office.tokens_excel().
            book.Save()
            book.Close(SaveChanges=False)
            book=None
            proof=app.Workbooks.Open(Filename=str(scratch),UpdateLinks=0,
                                     ReadOnly=True,AddToMru=False)
            try:verify_converted_three_layer_sheet(proof.Worksheets(1))
            finally:proof.Close(SaveChanges=False)
            if sha(source)!=original:
                raise RuntimeError('Исходный исторический акт изменился; операция остановлена.')
            if dest.exists():
                raise FileExistsError('Целевое имя занято перед публикацией.')
            os.rename(scratch,dest)
            return {'candidate':str(dest),'approved':False,
                    'source_sha256_unchanged':True,'candidate_sha256':sha(dest)}
        finally:
            if book is not None:
                try:book.Close(SaveChanges=False)
                except Exception:pass
            if app is not None:
                try:app.Quit()
                except Exception:pass
            pythoncom.CoUninitialize()
    finally:
        # Keep failed scratch copies for diagnosis, but never change source.
        pass


if __name__=='__main__':
    parser=argparse.ArgumentParser()
    parser.add_argument('source',help='Исторический трёхслойный акт, только чтение')
    parser.add_argument('destination',help='Новый утверждаемый XLSX')
    args=parser.parse_args()
    print(json.dumps(convert(args.source,args.destination),ensure_ascii=False,indent=2))
,'').upper()
        expected=col+'18:'+{'E':'G','H':'I','J':'K'}[col]+'18'
        if merge!=expected:
            raise ValueError('Не совпадает объединённая ячейка '+col+'18')
    return True


def verify_converted_three_layer_sheet(sheet):
    """Check complete template after native Excel Save/ReadOnly Open."""
    for col,index in LAYERS:
        if sheet.Range(col+'17').Value2!=index:
            raise ValueError('После сохранения потерян слой №'+str(index))
        prefix='AKZ_LAYER_'+str(index)+'_'
        fields=((18,'MATERIAL'),(19,'PAINT_BATCH'),(20,'HARDENER_BATCH'),
                (21,'COLOR'),(22,'APPLICATION_DATE'),(23,'TIME_START'),
                (24,'TIME_END'),(27,'WET_AVG'))
        for row,field in fields:
            if sheet.Range(col+str(row)).Value2!='{{'+prefix+field+'}}':
                raise ValueError('Потеряно поле '+col+str(row))
        for row,field in ((25,'RANGE'),(26,'TOTAL_RANGE'),(28,'DRY_AVG')):
            if sheet.Range(col+str(row)).Value2!='{{LAYER_'+str(index)+'_'+field+'}}':
                raise ValueError('Потеряна толщина/средняя слоя №'+str(index))
        for row,field in ((31,'SURFACE'),(32,'AIR'),(33,'HUMIDITY'),(34,'DEW')):
            if sheet.Range(col+str(row)).Value2!='{{'+field+'}}':
                raise ValueError('Потерян климат слоя №'+str(index))
    if sheet.Range('E7').Value2!='{{OBJECT}}':
        raise ValueError('Потеряно поле объекта АКЗ.')
    if not str(sheet.PageSetup.PrintArea or '').strip():
        raise ValueError('Потеряна область печати АКЗ.')
    return True


def convert(source, destination):
    source=Path(source).resolve(strict=True)
    dest=Path(destination).absolute()
    if source.suffix.lower()!='.xlsx' or dest.suffix.lower()!='.xlsx':
        raise ValueError('Требуется исходный и новый файл XLSX.')
    if dest.exists() or source==dest:
        raise FileExistsError('Новый шаблон должен иметь ещё не существующее имя.')
    if not dest.parent.is_dir():
        raise FileNotFoundError('Создайте отдельную папку для шаблона.')
    original=sha(source)
    scratch=dest.parent/('.vzmk-three-coat-'+uuid4().hex+'.xlsx')
    shutil.copyfile(source,scratch)
    app=None;book=None
    try:
        import pythoncom
        import win32com.client
        pythoncom.CoInitialize()
        try:
            app=win32com.client.DispatchEx('Excel.Application')
            app.Visible=False;app.DisplayAlerts=False;app.EnableEvents=False
            app.AskToUpdateLinks=False;app.AutomationSecurity=3
            book=app.Workbooks.Open(Filename=str(scratch),UpdateLinks=0,
                                    ReadOnly=False,AddToMru=False)
            if book.Worksheets.Count!=1 or book.ReadOnly:
                raise RuntimeError('Ожидался один редактируемый лист.')
            sheet=book.Worksheets(1)
            if sheet.ProtectContents:
                raise RuntimeError('Лист защищён.')
            if 'Слой №' not in str(sheet.Range('B17').Value2):
                raise RuntimeError('Неизвестная форма: не найдена таблица трёх слоёв.')
            if 'толщина' not in str(sheet.Range('C28').Value2 or '').casefold():
                raise RuntimeError('Неизвестная форма: нет строки средней сухой толщины.')

            # Only edit the fresh copy, leaving merges, dimensions, image,
            # print areas, page setup and all approved static labels in place.
            for cell,token in STATIC.items():
                sheet.Range(cell).Value2=token
            for column,n in LAYERS:
                pref='AKZ_LAYER_'+str(n)+'_'
                fields={
                    18:pref+'MATERIAL',
                    19:pref+'PAINT_BATCH',
                    20:pref+'HARDENER_BATCH',
                    21:pref+'COLOR',
                    22:pref+'APPLICATION_DATE',
                    23:pref+'TIME_START',
                    24:pref+'TIME_END',
                    25:'LAYER_'+str(n)+'_RANGE',
                    26:'LAYER_'+str(n)+'_TOTAL_RANGE',
                    27:pref+'WET_AVG',
                    28:'LAYER_'+str(n)+'_DRY_AVG',
                    31:'SURFACE',32:'AIR',33:'HUMIDITY',34:'DEW',
                }
                for row,field in fields.items():
                    cell=sheet.Range(column+str(row))
                    cell.Value2='{{'+field+'}}'
            # Original H4 and per-coat row 22 were Excel dates. They are
            # restored to true numbers by production_app.office.tokens_excel().
            # Existing E30/H30/J30 formulas link to the corresponding row 22.
            book.Save()
            book.Close(SaveChanges=False);book=None
            if sha(source)!=original:
                raise RuntimeError('Исходный исторический акт изменился; операция остановлена.')
            if dest.exists():
                raise FileExistsError('Целевое имя занято перед публикацией.')
            os.rename(scratch,dest)
            return {'candidate':str(dest),'approved':False,
                    'source_sha256_unchanged':True,'candidate_sha256':sha(dest)}
        finally:
            if book is not None:
                try:book.Close(SaveChanges=False)
                except Exception:pass
            if app is not None:
                try:app.Quit()
                except Exception:pass
            pythoncom.CoUninitialize()
    finally:
        # Keep failed scratch copies for diagnosis, but never change source.
        pass


if __name__=='__main__':
    parser=argparse.ArgumentParser()
    parser.add_argument('source',help='Исторический трёхслойный акт, только чтение')
    parser.add_argument('destination',help='Новый утверждаемый XLSX')
    args=parser.parse_args()
    print(json.dumps(convert(args.source,args.destination),ensure_ascii=False,indent=2))
