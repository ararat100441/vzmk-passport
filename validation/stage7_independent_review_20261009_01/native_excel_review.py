import sys
sys.dont_write_bytecode=True
import json,hashlib
from decimal import Decimal
from pathlib import Path
import pythoncom,win32com.client
OUT=Path(__file__).resolve().parent;ROOT=OUT.parents[1]
m=json.loads((ROOT/'validation/stage7_resume_20261009_175000/ASSEMBLY_01/STAGE7_SELECTED_TEST_MANIFEST.json').read_text(encoding='utf8'))
p=Path(m['membership']['vom_path'])
def sha(p):return hashlib.sha256(p.read_bytes()).hexdigest()
before=sha(p);pythoncom.CoInitialize();app=book=None
try:
    app=win32com.client.DispatchEx('Excel.Application');app.Visible=False;app.DisplayAlerts=False
    app.AutomationSecurity=3;app.EnableEvents=False;app.AskToUpdateLinks=False
    book=app.Workbooks.Open(str(p),UpdateLinks=0,ReadOnly=True,AddToMru=False,CorruptLoad=0)
    assert book.ReadOnly
    candidates=[]
    for s in book.Worksheets:
        if s.Visible!=-1:continue
        used=s.UsedRange;data=used.Value2
        for i,row in enumerate(data[:30]):
            labels=[str(v or '').strip().lower() for v in row]
            if 'отправочная марка' in labels and 'количество' in labels:
                mc=labels.index('отправочная марка');qc=labels.index('количество')
                massc=next(j for j,x in enumerate(labels) if x.startswith('масса') and 'кг' in x)+1
                assert str(data[i+1][massc]).strip().lower()=='всех'
                candidates.append((s,used.Row,data,i,mc,qc,massc))
    assert len(candidates)==1
    s,start,data,header,mc,qc,massc=candidates[0]
    visible=[];hidden=[]
    for i in range(header+2,len(data)):
        mark=data[i][mc]
        if mark is None or str(mark).strip().lower() in ('','итого','всего'):continue
        rn=start+i
        if s.Rows(rn).Hidden:hidden.append(rn)
        else:visible.append({'row':rn,'mark':mark,'quantity':int(data[i][qc]),'mass_kg':str(Decimal(str(data[i][massc])))})
    r={'path':str(p),'sha_before':before,'sheet':s.Name,'read_only':bool(book.ReadOnly),'normal_open':True,'repair':False,
        'visible_count':len(visible),'hidden_count':len(hidden),'quantity':sum(v['quantity'] for v in visible),
        'mass_kg':str(sum(Decimal(v['mass_kg']) for v in visible).quantize(Decimal('.01'))),'rows':visible,'hidden_rows':hidden,
        'external_links':book.LinkSources(1),'connections':book.Connections.Count,'version':app.Version,
        'macros_disabled':True,'save':False,'update_links':False}
finally:
    if book is not None:book.Close(SaveChanges=False)
    if app is not None:app.Quit()
    pythoncom.CoUninitialize()
r['sha_after']=sha(p);assert before==r['sha_after']
assert [r['visible_count'],r['hidden_count'],r['quantity'],r['mass_kg']]==[104,397,288,'15822.14']
pre=json.loads((ROOT/'validation/stage7_resume_20261009_175000/ASSEMBLY_01/PREFLIGHT.json').read_text(encoding='utf8'))
assert r['rows']==pre['preview']['vom']['rows']
r['implementation_preview_rows_equal']=True
(OUT/'NATIVE_EXCEL_VOM.json').write_text(json.dumps(r,ensure_ascii=False,indent=2),encoding='utf8')
print(json.dumps({k:v for k,v in r.items() if k not in ('rows','hidden_rows')},ensure_ascii=False))
