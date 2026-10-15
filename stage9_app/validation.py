"""Native read-only verification of a specific generated TEST bundle."""
from pathlib import Path
from uuid import uuid4
from decimal import Decimal
from .workspace import save
from .core import verify_manifest,document_groups,document_counts
from stage8_closure import storage as S

def inspect_bundle(root,path):
    path=S.no_links(path)
    if not path.is_relative_to(root/'validation'):raise S.Stop('MANIFEST_OUTSIDE_OWN_WORKSPACE')
    m=verify_manifest(path);target=root/'validation'/('native_'+uuid4().hex[:12]);target.mkdir()
    from passport_app.vom import read_vom
    appendix=next(r for r in m['records'] if r['type']=='QUALITY_APPENDIX_1')
    saved=read_vom(appendix['path']);visible=read_vom(path.parent/'VISIBLE_VOM.xlsx')
    if saved.quantity!=visible.quantity or saved.mass!=visible.mass or saved.area!=visible.area or saved.connections!=visible.connections or len(saved.rows)!=len(visible.rows):raise S.Stop('SAVED_APPENDIX_TOTALS_OR_M1_CHANGED')
    if any(a[:3]!=b[:3] or any(abs(x-y)>Decimal('.00000001') for x,y in zip(a[3:],b[3:])) for a,b in zip(saved.rows,visible.rows)):raise S.Stop('SAVED_APPENDIX_ROWS_CHANGED')
    native=[]
    from passport_app.vik import word_session,open_word
    from stage8_closure.office import excel_session
    with word_session() as word:
        for group in document_groups(m):
            if group['path'].suffix.lower()!='.docx':continue
            records=group['records'];r=records[0]
            before=S.sha(r['path']);doc=open_word(word,Path(r['path']),readonly=True)
            try:
                doc.Repaginate();text=doc.Content.Text;pages=doc.ComputeStatistics(2)
                required=[m['passport']]+[item[key] for item in records for key in ('number','date') if item.get(key)]
                if not doc.ReadOnly or pages<1 or any(t not in text for t in required):raise S.Stop('NATIVE_WORD_CONTENT_OR_READONLY')
                if r['type']=='QUALITY' and pages!=2:raise S.Stop('QUALITY_PAGES_CHANGED')
                if r['type']=='UZK' and 'РЕЗУЛЬТАТЫ КОНТРОЛЯ НЕ ПОДТВЕРЖДЕНЫ' not in text:raise S.Stop('UZK_SAMPLE_LABEL_LOST')
                native.append({'type':r['type'],'path':r['path'],'record_count':len(records),'record_numbers':[item['number'] for item in records],'pages':pages,'readonly':True,'text':text,'sha256':before})
            finally:doc.Close(False)
            if S.sha(r['path'])!=before:raise S.Stop('NATIVE_WORD_CHANGED_SOURCE')
    books=[]
    with excel_session() as excel:
        for group in document_groups(m):
            if group['path'].suffix.lower()!='.xlsx':continue
            r=group['records'][0]
            before=S.sha(r['path']);book=excel.Workbooks.Open(Filename=r['path'],UpdateLinks=0,ReadOnly=True,IgnoreReadOnlyRecommended=True,Notify=False,AddToMru=False,CorruptLoad=0)
            try:
                if not book.ReadOnly or book.Worksheets.Count<1:raise S.Stop('NATIVE_EXCEL_READONLY_OR_SHEETS')
                books.append({'type':r['type'],'path':r['path'],'readonly':True,'sheets':book.Worksheets.Count,'sha256':before})
            finally:book.Close(False)
            if S.sha(r['path'])!=before:raise S.Stop('NATIVE_EXCEL_CHANGED_SOURCE')
    result={'status':'PASS','manifest':str(path),'rows':len(saved.rows),'quantity':str(saved.quantity),'mass_kg':str(saved.mass),'area_m2':str(saved.area),'connections_m1_equal':True,**document_counts(m),'word':native,'excel':books}
    if len(native)!=result['word_documents'] or len(books)!=result['excel_documents']:raise S.Stop('NATIVE_PHYSICAL_FILE_COUNT_MISMATCH')
    save(target/'VALIDATION.json',result)
    return {k:v for k,v in result.items() if k not in ('word','excel','status')}|{'evidence':str(target/'VALIDATION.json')}
