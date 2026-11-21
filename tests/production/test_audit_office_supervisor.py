import json
import os
import sys
import tempfile
import unittest
from datetime import datetime, timezone
from pathlib import Path
from unittest.mock import Mock, patch

from stage9_app import supervisor


class SupervisorFailureTests(unittest.TestCase):
    def simulate(self, directory, mode, sharing=False):
        child = Mock(pid=41001,returncode=0)
        child.poll.return_value = None if mode=='hang' else 0
        child.wait.return_value = 1 if mode=='hang' else 0
        def launch(argv, **kwargs):
            if mode!='hang':
                path=Path(argv[argv.index('--result')+1])
                path.write_text('{' if mode=='corrupt' else '{"status":"PASS"}')
            return child
        secure_api = {n:Mock() for n in ['win32api','win32process','win32con','win32event','win32gui']}
        secure_api['win32process'].EnumProcesses.return_value=[]
        secure_api['win32process'].GetProcessTimes.return_value={'CreationTime':datetime.now(timezone.utc)}
        target=directory/'result.json'
        now=iter(range(1000))
        old_read=Path.read_text
        attempts=[]
        sleeps=[]
        def sleep(seconds):
            sleeps.append(seconds)
            if len(sleeps)>20:raise AssertionError('worker wait is unbounded')
        def simulated_read(path,*args,**kwargs):
            if path.name=='worker_result.json' and sharing:
                attempts.append(path)
                if len(attempts)==1:raise PermissionError('brief sharing violation')
            return old_read(path,*args,**kwargs)
        with (patch.dict(sys.modules,secure_api), patch.dict(os.environ,{'LOCALAPPDATA': str(directory),'VZMK_OFFICE_TASK_TIMEOUT_SECONDS':'1'}),
              patch.object(supervisor,'process_image',return_value='owned-exe'), patch.object(supervisor.subprocess,'Popen',side_effect=launch),
              patch.object(supervisor.time,'monotonic',side_effect=lambda:next(now)),patch.object(supervisor.time,'sleep',side_effect=sleep),
              patch.object(Path,'read_text',autospec=True,side_effect=simulated_read)):
            code=supervisor.supervise(['--task','readback','--result',str(target)])
        result=json.loads(target.read_text())
        summary=json.loads(Path(result['office_supervisor']).read_text())
        return code,result,summary,child

    def test_safe_task_deadline_finishes_with_failure_and_diagnostics(self):
        with tempfile.TemporaryDirectory() as folder:
            code,result,summary,worker=self.simulate(Path(folder),'hang')
        self.assertEqual((code,result['error']), (1,'OFFICE_WORKER_TIMEOUT'))
        self.assertTrue(summary['worker_timed_out'])
        worker.terminate.assert_called_once()

    def test_commit_and_recovery_have_no_forced_worker_deadline(self):
        for task in ['commit','complete','recover','compensate']:
            self.assertIsNone(supervisor._worker_timeout(['--task',task]))

    def test_brief_result_lock_preserves_success(self):
        with tempfile.TemporaryDirectory() as folder:
            code,result,summary,worker=self.simulate(Path(folder),'ok',sharing=True)
        self.assertEqual((code,result['status'],summary['ownership_errors']), (0,'PASS',[]))
        worker.terminate.assert_not_called()

    def test_bad_result_still_writes_supervisor_evidence(self):
        with tempfile.TemporaryDirectory() as folder:
            code,result,summary,worker=self.simulate(Path(folder),'corrupt')
        self.assertEqual((code,result['status']), (1,'FAIL'))
        self.assertTrue(summary['ownership_errors'])


if __name__ == '__main__':
    unittest.main()
