import sys
sys.dont_write_bytecode=True
import json, traceback
from pathlib import Path
import pythoncom, win32com.client
from audit_readonly import OUT, PROJECT, PASSPORT, TEMPLATES, allowed, sha, save

pythoncom.CoInitialize()
excel=None
results=[]
try:
    excel=win32com.client.DispatchEx('Excel.Application')
    excel.Visible=False
    excel.AutomationSecurity=3
    excel.EnableEvents=False
    excel.DisplayAlerts=False
    excel.AskToUpdateLinks=False
    try: excel.Calculation=-4135
    except Exception: pass
    try: printer=excel.ActivePrinter
    except Exception as e: printer=str(e)
    items=json.loads((OUT/'inventory_xlsx.json').read_text(encoding='utf-8'))
    chosen=[Path(x['path']) for x in items if 'АКЗ' in Path(x['path']).name or Path(x['path']).name.startswith('Акт')]
    chosen+= sorted(p for p in PASSPORT.rglob('*.xls') if 'ВОМ' in p.name)
    chosen+=[PROJECT/'ВОМ.xls']
    if len(sys.argv)>1 and sys.argv[1]=='extra':
        chosen=[Path(x['path']) for x in items if x['id']>=39]
    for i,p in enumerate(chosen,1):
        allowed(p)
        print('NATIVE',i,p.name,flush=True)
        rec={'path':str(p),'sha256_before':sha(p),'sheets':[]}
        book=None
        try:
            book=excel.Workbooks.Open(Filename=str(p),UpdateLinks=0,ReadOnly=True,
                IgnoreReadOnlyRecommended=True,AddToMru=False,Notify=False)
            rec['read_only']=bool(book.ReadOnly)
            assert rec['read_only']
            rec['file_format']=book.FileFormat
            rec['connections']=book.Connections.Count
            rec['macros_disabled']=excel.AutomationSecurity==3
            rec['events_disabled']=not excel.EnableEvents
            rec['calculation_mode_on_open']=excel.Calculation
            for si in range(1,book.Worksheets.Count+1):
                s=book.Worksheets(si)
                data={'name':s.Name,'visible':s.Visible,'used_range':s.UsedRange.Address,
                    'print_area':s.PageSetup.PrintArea,'orientation':s.PageSetup.Orientation,
                    'paper_size':s.PageSetup.PaperSize,'zoom':s.PageSetup.Zoom,
                    'fit_wide':s.PageSetup.FitToPagesWide,'fit_tall':s.PageSetup.FitToPagesTall}
                try:
                    s.Activate()
                    data['native_print_pages']=s.PageSetup.Pages.Count
                    data['h_page_breaks']=[s.HPageBreaks(j).Location.Row for j in range(1,s.HPageBreaks.Count+1)]
                    data['v_page_breaks']=[s.VPageBreaks(j).Location.Column for j in range(1,s.VPageBreaks.Count+1)]
                except Exception as e: data['pagination_error']=str(e)
                if p.suffix.lower()=='.xls':
                    ur=s.UsedRange
                    rr,cc=ur.Rows.Count,ur.Columns.Count
                    assert rr<=5000 and cc<=100,(rr,cc)
                    vals=ur.Value2; forms=ur.Formula
                    if not isinstance(vals,tuple): vals=((vals,),); forms=((forms,),)
                    cells=[]; lines=[]
                    for ri,row in enumerate(vals):
                        parts=[]
                        for ci,val in enumerate(row):
                            if val is None: continue
                            rownum=ur.Row+ri; colnum=ur.Column+ci
                            col=''; n=colnum
                            while n: n,m=divmod(n-1,26); col=chr(65+m)+col
                            addr=col+str(rownum)
                            f=forms[ri][ci]
                            cell={'cell':addr,'value':val}
                            if isinstance(f,str) and f.startswith('='): cell['formula']=f
                            cells.append(cell)
                            parts.append(addr+'='+str(val).replace('\n',' / '))
                        if parts: lines.append(' | '.join(parts))
                    data['cells']=cells
                    (OUT/f'native_xls_{i:03}_{si}.txt').write_text('\n'.join(lines),encoding='utf-8')
                rec['sheets'].append(data)
        except Exception as e:
            rec['error']=str(e); rec['traceback']=traceback.format_exc()
        finally:
            if book is not None: book.Close(SaveChanges=False)
        rec['sha256_after']=sha(p)
        assert rec['sha256_before']==rec['sha256_after'],str(p)
        results.append(rec)
        save('native_excel_extra.json' if len(sys.argv)>1 else 'native_excel.json',{'active_printer':printer,'method':'PageSetup.Pages.Count, no XLM/VBA or printing/export',
                                 'files':results})
finally:
    if excel is not None: excel.Quit()
    pythoncom.CoUninitialize()
print('NATIVE_COMPLETE',len(results),flush=True)
