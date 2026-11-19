"""Explicit artificial TEST data. Never convert existing TEST/SAMPLE receipts."""
from pathlib import Path
import shutil
from uuid import uuid4
from .office import excel_session
from .storage import ROOT, isolated_parent, frozen, sha, write, now, Stop, digest
from .native import inspect


def create(parent, *, uzk=True):
    parent = isolated_parent(parent)
    parent.mkdir(parents=True,exist_ok=True)
    root = parent/uuid4().hex
    root.mkdir(exist_ok=False)
    for name in ('registry','journal','backups','transactions','inputs'):
        (root/name).mkdir()
    source = ROOT/'DEV/УЗК_DEV.xlsm'
    registry = root/'registry/УЗК_TEST.xlsm'
    source_sha = sha(source)
    with frozen([source]):
        shutil.copyfile(source,registry)
        if sha(registry) != source_sha:
            raise Stop('FIXTURE_COPY_FAILED')
    project = 'STAGE8 ИСКУССТВЕННЫЙ ПРОЕКТ'
    with excel_session() as excel:
        book = excel.Workbooks.Open(Filename=str(registry),UpdateLinks=0,ReadOnly=False,
            IgnoreReadOnlyRecommended=True,Notify=False,AddToMru=False,CorruptLoad=0)
        try:
            sheet = book.Worksheets('УЗК')
            for row,mark,stock in [(24,'S8-A',12),(25,'S8-B',10),(26,'S8-C',8)]:
                sheet.Cells(row,1).Value2 = project
                sheet.Cells(row,2).Value2 = mark
                sheet.Cells(row,3).Value2 = stock
            formulas = book.Worksheets.Add(After=book.Worksheets(book.Worksheets.Count))
            formulas.Name = 'S8_FORMULAS'
            formulas.Cells(1,1).Formula = "=SUM('УЗК'!C24:C26)"
            formulas.Cells(2,1).Formula = '=2+3'
            formulas.Cells(3,1).Value2 = 'ИСКУССТВЕННЫЙ TEST — НЕ ДЛЯ ВЫДАЧИ'
            formulas.Range('A1:A3').Font.Bold = True
            formulas.Range('A1:A3').Font.Size = 13
            formulas.Range('A1:A2').NumberFormat = '0'
            formulas.Columns('A:A').ColumnWidth = 55
            formulas.Visible = 2  # veryHidden, retained across the close.
            excel.CalculateFull()
            book.Save()
        finally:
            book.Close(SaveChanges=False)
    native = inspect(registry)
    meta = {'schema':'STAGE8_SANDBOX_V1','sandbox_id':root.name,
        'created_at':now(),'mode':'ISOLATED_TEST','official_issue':False,
        'initial_registry_sha256':sha(registry),'source_copy':{'path':str(source),'sha256':source_sha},
        'warning':'Entire closure dataset is artificial; never eligible for official issue.',
        'native_seed':native}
    write(root/'SANDBOX.json',meta | {'seal':digest(meta)})
    identity = {'project':project,'passport':'8001/2099','shipment_id':'SYNTHETIC-SHIPMENT-001','test_only':True}
    quantities = [{'mark':'S8-A','quantity':3},{'mark':'S8-B','quantity':4}] if uzk else []
    def put(name,data):
        p = root/'inputs'/name
        write(p,identity | data)
        return {'path':str(p),'sha256':sha(p)}
    shipment = put('SHIPMENT.json',{'schema':'STAGE8_SYNTHETIC_SHIPMENT_V1','rows':[
        {'mark':'S8-A','quantity':5},{'mark':'S8-B','quantity':6},{'mark':'S8-C','quantity':1}]})
    requirements = put('REQUIREMENTS.json',{'schema':'STAGE8_REQUIREMENTS_V1',
        'uzk_required':uzk,'akz_required':True,'uzk':quantities})
    control = put('CONTROL.json',{'schema':'STAGE8_CONFIRMED_TEST_CONTROL_V1',
        'status':'CONFIRMED_TEST','all_required_control_confirmed':True,
        'uzk':[q | {'confirmed':True} for q in quantities]})
    kinds = ['QUALITY','APPENDIX_1','VIK','AKZ','AKZ_APPENDIX'] + (['UZK'] if uzk else [])
    documents = [put(kind+'.json',{'schema':'STAGE8_SYNTHETIC_DOCUMENT_V1','kind':kind,
        'status':'CONFIRMED_TEST','results_confirmed':True,'content':'ARTIFICIAL TEST PLACEHOLDER; NOT AN ISSUED DOCUMENT'}) | {'kind':kind} for kind in kinds]
    request = root/'inputs/CLOSE_REQUEST.json'
    write(request,identity | {'schema':'STAGE8_CLOSE_REQUEST_V1','mode':'ISOLATED_TEST',
        'official_issue':False,'shipment':shipment,'requirements':requirements,'control':control,'documents':documents})
    if sha(source) != source_sha:
        raise Stop('SOURCE_CHANGED_DURING_FIXTURE_CREATION')
    return {'sandbox':str(root),'request':str(request),'mode':'ISOLATED_TEST'}
