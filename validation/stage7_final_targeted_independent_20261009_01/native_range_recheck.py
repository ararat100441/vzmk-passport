"""Resolve ranges with native Find, avoiding plain-text offsets across MERGEFIELDs."""
import sys
sys.dont_write_bytecode=True
import pythoncom,win32com.client
from review_checks import *
OUTDIR=OUT/'native_range_recheck';OUTDIR.mkdir(exist_ok=False)
BASE=read(OUT/'NATIVE_REVIEW.json')
ACTS='Акт ВИК № 9/2026 от 29.09.2026 г., Акт ВИК № 9/2026-1 от 29.09.2026 г., Акт УЗК № 9/2026 от 01.10.2026 г., Сертификат качества антикоррозийной обработки № 9/2026 от 05.10.2026 г.'
STORAGE='Документы о качестве на металлопрокат, сварочные материалы и лакокрасочные материалы хранятся на предприятии'
def find(doc,needle,start=0):
    r=doc.Range(start,doc.Content.End)
    require=r.Find.Execute(FindText=needle,MatchCase=True,MatchWholeWord=False,MatchWildcards=False,Forward=True,Wrap=0,Format=False)
    assert require and r.Text==needle,(needle,r.Text)
    return r
pythoncom.CoInitialize();word=None
try:
    word=win32com.client.DispatchEx('Word.Application')
    word.Visible=False;word.DisplayAlerts=0;word.AutomationSecurity=3;word.Options.UpdateLinksAtOpen=False
    for label in ('first','new','manual'):
        path=Path(BASE[label]['opened_path']);before=sha(path)
        d=OUTDIR/label;d.mkdir(exist_ok=False)
        doc=word.Documents.Open(FileName=str(path),ReadOnly=True,AddToRecentFiles=False,ConfirmConversions=False,OpenAndRepair=False,NoEncodingDialog=True)
        try:
            assert doc.ReadOnly
            doc.Repaginate();doc.ActiveWindow.View.Type=3
            a=find(doc,'Согласно условиям договора на поставку конструкции защищены от коррозии:')
            b=find(doc,'Документ о качестве составлен на основании приемочных актов:')
            r=doc.Range(a.Start,b.Start)
            ranges={'point12':{'start':r.Start,'end':r.End,'text':r.Text,'page':r.Information(3),
                'top_pt':r.Information(6),'end_pt':doc.Range(r.End-1,r.End).Information(6),'style':str(r.Style)}}
            (d/'point12.emf').write_bytes(bytes(r.EnhMetaFileBits))
            if label!='first':
                p13=find(doc,ACTS,b.End);p14=find(doc,ACTS,p13.End);storage=find(doc,STORAGE,p14.End)
                for k,r in [('point13',p13),('point14',p14),('storage',storage)]:
                    ranges[k]={'start':r.Start,'end':r.End,'text':r.Text,'page':r.Information(3),'top_pt':r.Information(6),
                        'font':{n:getattr(r.Font,n) for n in ('Name','Size','Bold','Italic','Underline','Color','Subscript','Superscript')},
                        'paragraph':{n:getattr(r.ParagraphFormat,n) for n in ('Alignment','LeftIndent','RightIndent','FirstLineIndent','SpaceBefore','SpaceAfter','LineSpacing','LineSpacingRule','KeepTogether','KeepWithNext')},'style':str(r.Style)}
            save('native_range_recheck/'+label+'/RANGES.json',ranges)
            BASE[label]['ranges']=ranges;BASE[label]['range_method']='Microsoft Word Find; asserted exact native Range.Text'
        finally:doc.Close(SaveChanges=0)
        assert sha(path)==before
        print(label,ranges['point12']['start'],ranges['point12']['end'],ranges['point12']['page'],flush=True)
finally:
    if word is not None:word.Quit(SaveChanges=0)
    pythoncom.CoUninitialize()
save('NATIVE_VERIFIED_REVIEW.json',BASE)
