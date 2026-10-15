import sys
sys.dont_write_bytecode=True
import json,hashlib
from pathlib import Path
import pythoncom,win32com.client
OUT=Path(__file__).resolve().parent;a=json.loads((OUT/'OUTPUT_HASHES.json').read_text(encoding='utf-8-sig'))
results={};pythoncom.CoInitialize();word=None
try:
    word=win32com.client.DispatchEx('Word.Application');word.Visible=False;word.DisplayAlerts=0;word.AutomationSecurity=3;word.Options.UpdateLinksAtOpen=False
    for k in ('template','quality'):
        p=Path(a[k]['path']);before=hashlib.sha256(p.read_bytes()).hexdigest()
        doc=word.Documents.Open(FileName=str(p),ReadOnly=True,AddToRecentFiles=False,ConfirmConversions=False,OpenAndRepair=False,NoEncodingDialog=True)
        try:
            doc.Repaginate();doc.ActiveWindow.View.Type=3
            r=doc.Content.Duplicate;r.Find.ClearFormatting();r.Find.Text='толщина слоя';r.Find.Wrap=0
            assert r.Find.Execute()
            chars=[]
            for c in r.Characters:
                f=c.Font
                chars.append({'text':c.Text,'font':f.Name,'size':f.Size,'bold':f.Bold,'italic':f.Italic,'spacing':f.Spacing,
                    'scaling':f.Scaling,'position':f.Position,'kerning':f.Kerning,'language':c.LanguageID,
                    'x':c.Information(5),'y':c.Information(6),'page':c.Information(3)})
            results[k]=chars
        finally:doc.Close(SaveChanges=0)
        assert before==hashlib.sha256(p.read_bytes()).hexdigest()
finally:
    if word is not None:word.Quit(SaveChanges=0)
    pythoncom.CoUninitialize()
results['equal']=results['template']==results['quality']
(OUT/'POINT12_NATIVE_FONT_POSITIONS.json').write_text(json.dumps(results,ensure_ascii=False,indent=2),encoding='utf8')
print(json.dumps(results,ensure_ascii=False))
