"""Dedicated read-only viewer; never attaches to the user's Office instance."""
from pathlib import Path
from uuid import uuid4
from .diagnostics import viewer_diagnostic,event,error_record
class Viewer:
    def __init__(self):
        self.word=None;self.excel=None;self.closed=False
        self.cleanup_errors=[];self.close_result=None;self.open_attempt_id=None
        from .ownership import owned_office
        from .gui import config_path
        self.diagnostic=viewer_diagnostic(config_path().parent/'diagnostics')
        self.viewer_id=self.diagnostic.parent.name
        self.guard=owned_office(self.diagnostic,cleanup_errors=self.cleanup_errors)
        self.guard.__enter__()
    def open(self,paths):
        if self.closed:raise RuntimeError('VIEWER_ALREADY_CLOSED')
        import win32com.client
        from passport_app.vik import open_word
        self.open_attempt_id=uuid4().hex
        paths=list(dict.fromkeys(str(Path(p)) for p in paths))
        try:
            event(self.diagnostic,'OPEN_BEGIN',{'viewer_id':self.viewer_id,'open_attempt_id':self.open_attempt_id,'paths':paths})
            for path in paths:
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
            event(self.diagnostic,'OPENED',{'status':'PASS','viewer_id':self.viewer_id,'open_attempt_id':self.open_attempt_id})
        except Exception as primary:
            self.close(primary=primary)
            raise
    def close(self,primary=None):
        if self.closed:return self.close_result
        close_attempt_id=uuid4().hex
        def attempt(step,action):
            try:action()
            except Exception as exc:self.cleanup_errors.append(error_record(step,exc))
        attempt('Diagnostic.CloseBegin',lambda:event(self.diagnostic,'CLOSE_BEGIN',{
            'viewer_id':self.viewer_id,'open_attempt_id':self.open_attempt_id,'close_attempt_id':close_attempt_id}))
        for app,collection in ((self.word,'Documents'),(self.excel,'Workbooks')):
            if app is None:continue
            name='Word' if collection=='Documents' else 'Excel'
            if collection=='Workbooks':
                def hide():app.Visible=False;app.UserControl=False
                attempt('Excel.Hide',hide)
            attempt(name+'.Close',lambda:app.Documents.Close(False) if collection=='Documents' else app.Workbooks.Close())
            attempt(name+'.Quit',app.Quit)
        self.word=self.excel=None;app=None;self.closed=True
        attempt('Office.GuardExit',lambda:self.guard.__exit__(type(primary) if primary is not None else None,
                                                           primary,primary.__traceback__ if primary is not None else None))
        self.close_result={'status':'FAIL' if primary is not None or self.cleanup_errors else 'PASS',
            'viewer_id':self.viewer_id,'open_attempt_id':self.open_attempt_id,'close_attempt_id':close_attempt_id,
            'diagnostic':str(self.diagnostic),'cleanup_errors':self.cleanup_errors,
            'primary_error':error_record('Primary',primary) if primary is not None else None}
        attempt('Diagnostic.CloseResult',lambda:event(self.diagnostic,'CLOSE_RESULT',self.close_result))
        if self.cleanup_errors:self.close_result['status']='FAIL'
        if primary is not None and self.cleanup_errors:primary.add_note('Viewer cleanup: '+str(self.close_result))
        return self.close_result
