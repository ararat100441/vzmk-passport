from common import *
import win32com.client,fitz
from stage9_app.ownership import owned_office
assert not (OUT/'FOREIGN_OFFICE_STOP.json').exists()
rows=read(OUT/'INDEPENDENT_XLSX_AUDIT.json')['files'];selected=[r for r in rows if r['set']=='FIRST']
folder=OUT/'native_print';folder.mkdir();results=[];excel=book=None
with owned_office(folder/'ownership.json') as owned:
 try:
  excel=win32com.client.DispatchEx('Excel.Application');excel.AutomationSecurity=3;excel.EnableEvents=False;excel.DisplayAlerts=False;excel.AskToUpdateLinks=False
  for row in selected:
   p=Path(row['path']);before=sha(p);book=excel.Workbooks.Open(Filename=str(p),UpdateLinks=0,ReadOnly=True,IgnoreReadOnlyRecommended=True,AddToMru=False,CorruptLoad=0,Notify=False)
   assert book.ReadOnly;excel.CalculateFull()
   sheets=[]
   for sheet in book.Worksheets:
    sheets.append({'name':sheet.Name,'visible':sheet.Visible,'print_area':sheet.PageSetup.PrintArea,'pages':int(sheet.PageSetup.Pages.Count),'center_header':sheet.PageSetup.CenterHeader,'center_footer':sheet.PageSetup.CenterFooter,'fit_wide':sheet.PageSetup.FitToPagesWide,'fit_tall':sheet.PageSetup.FitToPagesTall})
   target=folder/('AKZ'+row['variant']+'_'+row['kind']+'.pdf');book.ExportAsFixedFormat(0,str(target),0,True,False);book.Close(SaveChanges=False);book=None
   assert sha(p)==before
   pdf=fitz.open(target);images=[];texts=[]
   for index,page in enumerate(pdf):
    text=page.get_text();texts.append(text);image=target.with_name(target.stem+f'_page{index+1}.png');page.get_pixmap(matrix=fitz.Matrix(1.5,1.5)).save(image);images.append(str(image))
   results.append({'file':str(p),'before_after_SHA_equal':True,'read_only':True,'normal_load_CorruptLoad':0,'sheets':sheets,'pdf':str(target),'pdf_sha256':sha(target),'pages':len(pdf),'images':images,'texts':texts});pdf.close()
 finally:
  if book is not None:book.Close(SaveChanges=False)
  if excel is not None:excel.Quit()
  sheet=book=excel=None
assert not any(r.get('forced_termination_attempted') for r in read(folder/'ownership.json')['instances'])
save(OUT/'NATIVE_EXCEL_PRINT.json',{'status':'PASS_NATIVE_EXPORT','files':results,'printed_to_physical_printer':False,'visual_inspection':'SEPARATE_MODEL_REVIEW','GUI':False,'ownership':str(folder/'ownership.json')})
print('Native Excel readonly/print export PASS',len(results),'books',sum(x['pages'] for x in results),'pages',flush=True)
