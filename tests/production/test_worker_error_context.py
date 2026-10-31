"""A cleanup failure must not erase the failure which caused cleanup."""
import json,tempfile,unittest
from pathlib import Path
from contextlib import contextmanager
from unittest.mock import patch
import production

class WorkerErrorContext(unittest.TestCase):
    def test_primary_failure_survives_failed_office_teardown(self):
        @contextmanager
        def failed_cleanup(*args,**kwargs):
            yield
            raise RuntimeError('CLEANUP_FAILURE')
        with tempfile.TemporaryDirectory() as folder:
            root=Path(folder);request=root/'request.json';result=root/'result.json'
            request.write_text('{}',encoding='utf8')
            argv=['production.py','--task','check','--request',str(request),
                  '--result',str(result),'--office-worker',str(root/'office.json')]
            with patch('sys.argv',argv),patch('production.task',side_effect=RuntimeError('PRIMARY_NATIVEOM_FAILURE')),patch('stage9_app.ownership.owned_office',failed_cleanup):
                self.assertEqual(production.main(),1)
            actual=json.loads(result.read_text(encoding='utf8'))
            self.assertEqual(actual['status'],'FAIL')
            self.assertIn('PRIMARY_NATIVEOM_FAILURE',actual['error'])
            self.assertIn('CLEANUP_FAILURE',actual['cleanup_error'])
            self.assertIn('PRIMARY_NATIVEOM_FAILURE',actual['traceback'])

if __name__=='__main__':unittest.main()
