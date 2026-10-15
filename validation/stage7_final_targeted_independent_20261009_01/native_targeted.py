"""Owned read-only Microsoft Word session; no saves, repair, PDF, or generators."""
import sys
sys.dont_write_bytecode=True
import json
from pathlib import Path
import pythoncom,win32com.client
from review_checks import OUT,ROOT,NEW,FIRST,MANUAL,sha,save,read

results={}
pythoncom.CoInitialize()
word=None
try:
    word=win32com.client.DispatchEx('Word.Application')
    word.Visible=False;word.DisplayAlerts=0;word.AutomationSecurity=3
    word.Options.UpdateLinksAtOpen=False
    for label,source in [('new',NEW),('first',FIRST),('manual',MANUAL)]:
        before=sha(source)
        folder=OUT/('native_'+label);folder.mkdir(exist_ok=False)
        path=source
        # Isolate the user's already-open references from our own Word process.
        if label!='new':
            path=folder/'SHA_VERIFIED_READONLY_REFERENCE.docx'
            path.write_bytes(source.read_bytes());assert sha(path)==before
        doc=None
        try:
            doc=word.Documents.Open(FileName=str(path),ReadOnly=True,AddToRecentFiles=False,
                ConfirmConversions=False,OpenAndRepair=False,NoEncodingDialog=True)
            assert doc.ReadOnly
            doc.Repaginate();doc.ActiveWindow.View.Type=3
            n=doc.ComputeStatistics(2)
            result={'source':str(source),'opened_path':str(path),'sha_before':before,'ReadOnly':bool(doc.ReadOnly),
                'OpenAndRepair':False,'Save':False,'PDF':False,'word_version':word.Version,'word_build':word.Build,
                'pages':n,'sections':doc.Sections.Count,'tables':doc.Tables.Count,'width_pt':doc.PageSetup.PageWidth,'height_pt':doc.PageSetup.PageHeight,
                'shapes':doc.Shapes.Count,'inline_shapes':doc.InlineShapes.Count,'revisions':doc.Revisions.Count,
                'page_texts':[],'lists':[],'fields':[],'ranges':{}}
            for i in range(1,n+1):
                s=doc.GoTo(1,1,i).Start;e=doc.GoTo(1,1,i+1).Start if i<n else doc.Content.End
                result['page_texts'].append(doc.Range(s,e).Text)
                (folder/f'page-{i}.emf').write_bytes(bytes(doc.ActiveWindow.Panes(1).Pages(i).EnhMetaFileBits))
            for i in range(1,doc.Paragraphs.Count+1):
                p=doc.Paragraphs(i);labeltext=p.Range.ListFormat.ListString
                if labeltext:result['lists'].append({'label':labeltext,'text':p.Range.Text,'page':p.Range.Information(3)})
            for i in range(1,doc.Fields.Count+1):
                f=doc.Fields(i);result['fields'].append({'code':f.Code.Text,'result':f.Result.Text})
            raw=doc.Content.Text
            heading12='Согласно условиям договора на поставку конструкции защищены от коррозии:'
            heading13='Документ о качестве составлен на основании приемочных актов:'
            s=raw.index(heading12);e=raw.index(heading13,s+len(heading12))
            r=doc.Range(s,e)
            endpoint=doc.Range(e-1,e)
            result['ranges']['point12']={'start':s,'end':e,'text':r.Text,'page':r.Information(3),
                'top_pt':r.Information(6),'end_pt':endpoint.Information(6),'style':str(r.Style)}
            (folder/'point12.emf').write_bytes(bytes(r.EnhMetaFileBits))
            acts='Акт ВИК № 9/2026 от 29.09.2026 г., Акт ВИК № 9/2026-1 от 29.09.2026 г., Акт УЗК № 9/2026 от 01.10.2026 г., Сертификат качества антикоррозийной обработки № 9/2026 от 05.10.2026 г.'
            storage='Документы о качестве на металлопрокат, сварочные материалы и лакокрасочные материалы хранятся на предприятии'
            for k,needle,at in [('point13',acts,0),('point14',acts,raw.index(heading13)+len(heading13)+len(acts)),('storage',storage,0)]:
                if label=='first':continue
                s=raw.index(needle,at);e=s+len(needle);r=doc.Range(s,e)
                result['ranges'][k]={'start':s,'end':e,'text':r.Text,'page':r.Information(3),'top_pt':r.Information(6),
                    'font':{n:getattr(r.Font,n) for n in ('Name','Size','Bold','Italic','Underline','Color','Subscript','Superscript')},
                    'paragraph':{n:getattr(r.ParagraphFormat,n) for n in ('Alignment','LeftIndent','RightIndent','FirstLineIndent','SpaceBefore','SpaceAfter','LineSpacing','LineSpacingRule','KeepTogether','KeepWithNext')},
                    'style':str(r.Style)}
        finally:
            if doc is not None:doc.Close(SaveChanges=0)
        result['sha_after']=sha(source);result['opened_sha_after']=sha(path)
        assert result['sha_after']==result['opened_sha_after']==before
        save('native_'+label+'/WORD.json',result)
        results[label]=result
        print(label,n,result['sections'],result['tables'],result['sha_after'],flush=True)
finally:
    if word is not None:word.Quit(SaveChanges=0)
    pythoncom.CoUninitialize()
save('NATIVE_REVIEW.json',results)
