"""Independent content/native checks of saved files from the final EXE run."""
import sys,json,traceback
from pathlib import Path
from decimal import Decimal
ROOT=Path(__file__).resolve().parents[2];sys.path.insert(0,str(ROOT))
from stage9_app.workspace import bind,save
from stage9_app.ownership import owned_office
from stage8_closure.storage import read,sha
from stage9_app.core import verify_manifest
from passport_app.vom import read_vom
from stage7_quality.native import inspect
from stage8_closure.office import excel_session
OUT=ROOT/'validation/stage9_20261012_01'
summary=read(OUT/'EXE_FINAL_RESULT.json');run=Path(summary['evidence']).parent
meta=read(summary['evidence']);root,_=bind(summary['workspace'])
results=[]
with owned_office(OUT/'OUTPUTS_OFFICE.json'):
    for key in ('manifest','new_passport_manifest'):
        m=verify_manifest(meta[key]);target=OUT/('native_'+m['passport'].replace('/','_'));target.mkdir()
        appendix=next(r['path'] for r in m['records'] if r['type']=='QUALITY_APPENDIX_1')
        saved=read_vom(appendix);assert saved.quantity==288 and saved.mass==Decimal('15822.14') and saved.area==Decimal('1913.494')
        visible=read_vom(Path(meta[key]).parent/'VISIBLE_VOM.xlsx')
        assert len(saved.rows)==104 and saved.connections==visible.connections
        assert len(saved.rows)==len(visible.rows)
        for a,b in zip(saved.rows,visible.rows):
            assert a[:3]==b[:3] and all(abs(x-y)<Decimal('.00000001') for x,y in zip(a[3:],b[3:]))
        native=[]
        for kind in ('VIK','UZK','QUALITY'):
            record=next(r for r in m['records'] if r['type']==kind)
            required=[m['passport'],record['date']]
            if kind=='UZK':required+=['РЕЗУЛЬТАТЫ КОНТРОЛЯ НЕ ПОДТВЕРЖДЕНЫ','ДНО','годен']
            native.append(inspect(record['path'],target,kind=kind,required=required))
        with excel_session(diagnostics=target/'excel') as excel:
            for record in m['records']:
                if Path(record['path']).suffix!='.xlsx':continue
                before=sha(record['path'])
                b=excel.Workbooks.Open(Filename=record['path'],UpdateLinks=0,ReadOnly=True,IgnoreReadOnlyRecommended=True,Notify=False,AddToMru=False,CorruptLoad=0)
                try:assert b.ReadOnly and b.Worksheets.Count>=1
                finally:b.Close(False)
                assert sha(record['path'])==before
        results.append({'passport':m['passport'],'status':'PASS','rows':104,'quantity':str(saved.quantity),'mass_kg':str(saved.mass),'area_m2':str(saved.area),'connections_m1_equal':True,'all_original_rows_equal':True,'native_word':native,'native_excel_normal_readonly':'PASS'})
save(OUT/'OUTPUTS_VALIDATION.json',{'status':'PASS','exe_sha256':summary['exe_sha256'],'cases':results})
