import json,hashlib,sys
from pathlib import Path
import pythoncom,win32com.client
OUT=Path(__file__).resolve().parent
def sha(p):return hashlib.sha256(p.read_bytes()).hexdigest()
def save(n,x): (OUT/n).write_text(json.dumps(x,ensure_ascii=False,indent=2,default=str),encoding='utf8')
def get(obj,attr):
    try:return getattr(obj,attr)
    except Exception as e:return {'error':str(e)}
selected=json.loads((OUT/'selected_quality.json').read_text(encoding='utf8'))
paths=[Path(selected['path'])]
if len(sys.argv)>1:paths=[Path(p).resolve() for p in sys.argv[1:]]
pythoncom.CoInitialize();word=None
try:
    word=win32com.client.DispatchEx('Word.Application');word.Visible=False;word.DisplayAlerts=0;word.AutomationSecurity=3
    word.Options.UpdateLinksAtOpen=False
    for k,path in enumerate(paths,start=1 if len(sys.argv)>1 else 0):
        before=sha(path);doc=None
        try:
            doc=word.Documents.Open(FileName=str(path),ReadOnly=True,AddToRecentFiles=False,ConfirmConversions=False,OpenAndRepair=False,NoEncodingDialog=True)
            doc.ActiveWindow.View.Type=3;doc.Repaginate()
            count=doc.ComputeStatistics(2)
            d={'path':str(path),'sha_before':before,'readonly':doc.ReadOnly,'saved_on_open':doc.Saved,'pages_count':count,'sections_count':doc.Sections.Count,'tables_count':doc.Tables.Count,'fields_count':doc.Fields.Count,'inline_shapes_count':doc.InlineShapes.Count,'shapes_count':doc.Shapes.Count,'content_controls_count':doc.ContentControls.Count,'revisions_count':doc.Revisions.Count,'track_revisions':doc.TrackRevisions,'protection':doc.ProtectionType,'word_version':word.Version,'active_printer':word.ActivePrinter,'pages':[],'sections':[],'tables':[],'paragraphs':[],'fields':[],'inline_shapes':[],'shapes':[]}
            for i in range(1,count+1):
                a=doc.GoTo(What=1,Which=1,Count=i).Start
                b=doc.GoTo(What=1,Which=1,Count=i+1).Start if i<count else doc.Content.End
                r=doc.Range(a,b)
                p={'page':i,'start':a,'end':b,'text':r.Text,'tables_count':r.Tables.Count,'sections_count':r.Sections.Count,'text_sha256':hashlib.sha256(r.Text.encode('utf8')).hexdigest()}
                try:
                    bits=doc.ActiveWindow.Panes(1).Pages(i).EnhMetaFileBits
                    (OUT/f'word_{k+1}_page_{i:02}.emf').write_bytes(bytes(bits));p['native_page_emf']=f'word_{k+1}_page_{i:02}.emf'
                except Exception as e:p['native_page_image_error']=str(e)
                d['pages'].append(p)
            for i in range(1,doc.Sections.Count+1):
                s=doc.Sections(i);p=s.PageSetup
                e={'index':i,'start':s.Range.Start,'end':s.Range.End,'page_start':s.Range.Information(3),'setup':{x:get(p,x) for x in ('SectionStart','Orientation','PageWidth','PageHeight','TopMargin','BottomMargin','LeftMargin','RightMargin','HeaderDistance','FooterDistance','DifferentFirstPageHeaderFooter','OddAndEvenPagesHeaderFooter')},'headers':[],'footers':[]}
                for name,col in [('headers',s.Headers),('footers',s.Footers)]:
                    for j in range(1,col.Count+1):
                        z=col(j);e[name].append({'index':j,'exists':z.Exists,'link_previous':z.LinkToPrevious,'text':z.Range.Text,'fields':[{'code':z.Range.Fields(f).Code.Text,'result':z.Range.Fields(f).Result.Text} for f in range(1,z.Range.Fields.Count+1)],'shapes_count':z.Shapes.Count})
                d['sections'].append(e)
            for i in range(1,doc.Tables.Count+1):
                t=doc.Tables(i);e={'index':i,'start':t.Range.Start,'end':t.Range.End,'rows':get(t.Rows,'Count'),'columns':get(t.Columns,'Count'),'uniform':t.Uniform,'cells':[]}
                for c in t.Range.Cells:
                    e['cells'].append({'row':c.RowIndex,'column':c.ColumnIndex,'start':c.Range.Start,'end':c.Range.End,'page':c.Range.Information(3),'text':c.Range.Text,'width':c.Width})
                d['tables'].append(e)
            for i in range(1,doc.Paragraphs.Count+1):
                p=doc.Paragraphs(i);r=p.Range
                d['paragraphs'].append({'index':i,'start':r.Start,'end':r.End,'page':r.Information(3),'text':r.Text,'page_break_before':p.Format.PageBreakBefore,'keep_with_next':p.Format.KeepWithNext,'keep_together':p.Format.KeepTogether,'style':str(p.Style)})
            for i in range(1,doc.Fields.Count+1):
                f=doc.Fields(i);d['fields'].append({'index':i,'code':f.Code.Text,'result':f.Result.Text,'start':f.Code.Start})
            for i in range(1,doc.InlineShapes.Count+1):
                s=doc.InlineShapes(i);d['inline_shapes'].append({'index':i,'type':s.Type,'start':s.Range.Start,'page':s.Range.Information(3),'width':s.Width,'height':s.Height,'alt':s.AlternativeText})
            for i in range(1,doc.Shapes.Count+1):
                s=doc.Shapes(i);d['shapes'].append({'index':i,'type':s.Type,'anchor':s.Anchor.Start,'page':s.Anchor.Information(3),'width':s.Width,'height':s.Height,'name':s.Name,'alt':s.AlternativeText})
            d['saved_before_close']=doc.Saved
            doc.Close(SaveChanges=0);doc=None
            d['sha_after']=sha(path);d['integrity_pass']=d['sha_after']==before
            save(f'word_{k+1}.json',d)
            (OUT/f'word_{k+1}_pages.txt').write_text('\n\n'.join(f'PAGE {p["page"]} [{p["start"]},{p["end"]})\n{p["text"]}' for p in d['pages']),encoding='utf8')
            print(json.dumps({key:d[key] for key in ('path','pages_count','sections_count','tables_count','fields_count','inline_shapes_count','shapes_count','integrity_pass')},ensure_ascii=False))
        finally:
            if doc is not None:doc.Close(SaveChanges=0)
finally:
    if word is not None:word.Quit(SaveChanges=0)
    pythoncom.CoUninitialize()
