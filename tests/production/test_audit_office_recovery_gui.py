import tempfile
import unittest
from pathlib import Path
from unittest.mock import Mock, patch

from production_app import gui


class RecoveryChoiceTests(unittest.TestCase):
    def setUp(self):
        self.app = Mock(spec=gui.App)
        self.app.active = None
        self.app.cfg = {'registry': 'registry.xlsx'}
        self.app.root = Mock()
        self.app.result_path = 'readback-result.json'
        self.states = {
            Path('A/journal.json'): {'phase': 'DOCUMENTS_PUBLISHED', 'target': 'Первый объект'},
            Path('B/journal.json'): {'phase': 'NEEDS_READBACK', 'target': 'Выход', 'project': 'Второй объект', 'number': '2/2026'},
        }

    def recovery(self, paths, selection=1, resume=True, undo=False):
        with (patch.object(gui.core, 'pending', return_value=paths),
              patch.object(gui.core, 'journal_read', side_effect=self.states.__getitem__),
              patch.object(gui.simpledialog, 'askinteger', return_value=selection) as select,
              patch.object(gui.messagebox, 'askyesnocancel', return_value=resume) as prompt,
              patch.object(gui.messagebox, 'askyesno', return_value=undo)):
            gui.App.recovery(self.app)
        self.app.error.assert_not_called()
        if len(paths)>1:
            self.assertIn('Второй объект · паспорт №2/2026 (NEEDS_READBACK)', select.call_args.args[1])
        if prompt.called:
            self.assertIn('Второй объект · паспорт №2/2026', prompt.call_args.args[1])

    def test_retry_then_complete_uses_selected_journal_and_supervised_result(self):
        self.recovery(list(self.states), selection=2)
        task, req, passed = self.app.run.call_args.args
        self.assertEqual((task, req), ('readback', {'journal':str(Path('B/journal.json'))}))
        self.app.run.reset_mock()
        passed({'status':'PASS'})
        task, req, finish = self.app.run.call_args.args
        self.assertEqual((task, req), ('complete', {'journal':str(Path('B/journal.json')), 'readback_result':self.app.result_path}))
        finish({'status':'PASS'})
        self.app.refresh.assert_called_once()

    def test_unfinished_journal_and_explicit_undo(self):
        path = next(iter(self.states))
        self.recovery([path], undo=True)
        self.assertEqual(self.app.run.call_args.args[:2], ('recover', {'journal':str(path)}))

    def test_cancel_does_not_change_any_journal(self):
        self.recovery(list(self.states), selection=None)
        self.app.run.assert_not_called()
        self.recovery([Path('B/journal.json')], resume=False, undo=False)
        self.app.run.assert_not_called()

    def test_locked_finished_result_retries_same_result_then_advances_once(self):
        self.app.active=Mock(returncode=0)
        self.app.active.poll.return_value=0
        self.app.buttons=[]
        for widget in ['projects','numbers','new_widget','vom_widget','source_numbers','shipment_entry','new','progress']:
            setattr(self.app,widget,Mock())
        self.app.poll.side_effect=lambda *args:gui.App.poll(self.app,*args)
        callback=Mock()
        with patch.object(gui,'read',side_effect=[PermissionError('sharing'),{'status':'PASS'}]) as read:
            gui.App.poll(self.app,Path('job'),'commit',callback,[])
            callback.assert_not_called()
            self.app.root.after.assert_called_once()
            self.app.root.after.call_args.args[1]()
        self.assertEqual([c.args[0] for c in read.call_args_list],[Path('job/result.json')]*2)
        callback.assert_called_once_with({'status':'PASS'})
        self.app.error.assert_not_called()
        self.app.run.assert_not_called()

    def test_failed_amendment_readback_keeps_journal_for_operator(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            journal = root / 'journal.json'
            journal.write_text('{}')
            gui.save(root/'request.json', {'journal':str(journal)})
            gui.save(root/'result.json', {'status':'FAIL','error':'Word unavailable'})
            self.app.active = Mock(returncode=1)
            self.app.active.poll.return_value = 1
            for widget in ['projects','numbers','new_widget','vom_widget','source_numbers','shipment_entry','new','progress']:
                setattr(self.app, widget, Mock())
            self.app.adding_marks.return_value = True
            gui.App.poll(self.app,root,'readback',Mock(),[])
            self.app.error.assert_called_once()
            self.app.run.assert_not_called()


if __name__ == '__main__':
    unittest.main()
