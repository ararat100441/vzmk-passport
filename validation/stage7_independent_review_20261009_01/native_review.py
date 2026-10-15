"""Independent COM reader. No source save, no PDF, no product renderer."""
import sys
sys.dont_write_bytecode=True
import json,hashlib,subprocess
from pathlib import Path
import pythoncom, win32com.client

OUT=Path(__file__).resolve().parent
def sha(p):return hashlib.sha256(Path(p).read_bytes()).hexdigest()
a=json.loads((OUT/'OUTPUT_HASHES.json').read_text(encoding='utf-8-sig'))
names=['quality','template','vik','uzk']
if len(sys.argv)>1:
    OUT=OUT/'native_recheck';OUT.mkdir(exist_ok=False);names=['template','quality']
results={}
pythoncom.CoInitialize()
word=None
try:
    word=win32com.client.DispatchEx('Word.Application')
    word.Visible=False;word.DisplayAlerts=0;word.AutomationSecurity=3
    word.Options.UpdateLinksAtOpen=False
    for k in names:
        path=Path(a[k]['path']);before=sha(path)
        d=OUT/('native_'+k);d.mkdir(exist_ok=False)
        doc=word.Documents.Open(FileName=str(path),ReadOnly=True,AddToRecentFiles=False,
            ConfirmConversions=False,OpenAndRepair=False,NoEncodingDialog=True)
        try:
            doc.Repaginate();doc.ActiveWindow.View.Type=3
            n=doc.ComputeStatistics(2)
            r={'path':str(path),'sha_before':before,'ReadOnly':bool(doc.ReadOnly),'pages':n,
                'sections':doc.Sections.Count,'tables':doc.Tables.Count,'protection':doc.ProtectionType,
                'revisions':doc.Revisions.Count,'fields':[], 'lists':[], 'page_texts':[],
                'width':doc.PageSetup.PageWidth,'height':doc.PageSetup.PageHeight,
                'version':word.Version,'build':word.Build,'printer':word.ActivePrinter,
                'hyperlinks':doc.Hyperlinks.Count,'inline_shapes':doc.InlineShapes.Count,'shapes':doc.Shapes.Count,
                'open_and_repair':False,'save':False,'pdf_export':False}
            for i in range(1,n+1):
                s=doc.GoTo(1,1,i).Start;e=doc.GoTo(1,1,i+1).Start if i<n else doc.Content.End
                r['page_texts'].append(doc.Range(s,e).Text)
                (d/f'page-{i}.emf').write_bytes(bytes(doc.ActiveWindow.Panes(1).Pages(i).EnhMetaFileBits))
            for i in range(1,doc.Paragraphs.Count+1):
                p=doc.Paragraphs(i);label=p.Range.ListFormat.ListString
                if label:r['lists'].append({'label':label,'text':p.Range.Text,'page':p.Range.Information(3)})
            for i in range(1,doc.Fields.Count+1):
                f=doc.Fields(i);r['fields'].append({'code':f.Code.Text,'result':f.Result.Text})
        finally:doc.Close(SaveChanges=0)
        r['sha_after']=sha(path);assert r['sha_after']==before
        (d/'WORD.json').write_text(json.dumps(r,ensure_ascii=False,indent=2),encoding='utf8')
        results[k]=r
        print(k,n,r['sections'],r['tables'],before,flush=True)
finally:
    if word is not None:word.Quit(SaveChanges=0)
    pythoncom.CoUninitialize()
(OUT/'NATIVE_REVIEW.json').write_text(json.dumps(results,ensure_ascii=False,indent=2),encoding='utf8')
