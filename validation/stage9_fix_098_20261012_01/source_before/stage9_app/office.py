"""Dedicated read-only viewer; never attaches to the user's Office instance."""
from pathlib import Path
class Viewer:
    def __init__(self):
        self.word=None;self.excel=None;self.closed=False
        from .ownership import owned_office
        from .gui import config_path
        from uuid import uuid4
        self.guard=owned_office(config_path().parent/'diagnostics'/('viewer_'+uuid4().hex+'.json'));self.guard.__enter__()
    def open(self,paths):
        import win32com.client
        from passport_app.vik import open_word
        try:
            for path in dict.fromkeys(paths):
                path=Path(path)
                if path.suffix.lower()=='.docx':
                    if self.word is None:
                        self.word=win32com.client.DispatchEx('Word.Application');self.word.AutomationSecurity=3;self.word.Options.UpdateLinksAtOpen=False
                    open_word(self.word,path,readonly=True)
                else:
                    if self.excel is None:
                        self.excel=win32com.client.DispatchEx('Excel.Application');self.excel.AutomationSecurity=3;self.excel.EnableEvents=False;self.excel.AskToUpdateLinks=False;self.excel.DisplayAlerts=False
                    self.excel.Workbooks.Open(Filename=str(path),UpdateLinks=0,ReadOnly=True,IgnoreReadOnlyRecommended=True,Notify=False,AddToMru=False,CorruptLoad=0)
            if self.word is not None:self.word.Visible=True
            if self.excel is not None:self.excel.Visible=True
        except Exception:self.close();raise
    def close(self):
        if self.closed:return []
        errors=[]
        for app,collection in ((self.word,'Documents'),(self.excel,'Workbooks')):
            if app is None:continue
            if collection=='Workbooks':
                try:app.Visible=False;app.UserControl=False
                except Exception as exc:errors.append(str(exc))
            try:
                if collection=='Documents':app.Documents.Close(False)
                else:app.Workbooks.Close()
            except Exception as exc:errors.append(str(exc))
            try:app.Quit()
            except Exception as exc:errors.append(str(exc))
        self.word=self.excel=None
        app=None
        self.closed=True;self.guard.__exit__(None,None,None)
        return errors
