"""Execute the real supervisor's final verdict AST against refusal scenarios.

No Win32 process is addressed by these tests. Native ownership and helper use
are covered separately by native_contract.py.
"""
import ast,copy,json,sys,tempfile,unittest
from pathlib import Path
from unittest.mock import patch
ROOT=Path(__file__).resolve().parents[2];sys.path.insert(0,str(ROOT))
from stage9_app.supervisor import cleanup_entry

TREE=ast.parse((ROOT/'stage9_app/supervisor.py').read_text(encoding='utf-8-sig'))
FUNCTION=next(n for n in TREE.body if isinstance(n,ast.FunctionDef) and n.name=='supervise')
ASSIGN=next(n for n in ast.walk(FUNCTION) if isinstance(n,ast.Assign) and any(isinstance(t,ast.Name) and t.id=='unbound_cleanup_failed' for t in n.targets))
VERDICT=next(n for n in ast.walk(FUNCTION) if isinstance(n,ast.If) and isinstance(n.test,ast.BoolOp)
             and any(isinstance(v,ast.Name) and v.id=='ownership_errors' for v in n.test.values))
PROGRAM=compile(ast.fix_missing_locations(ast.Module(body=[ASSIGN,VERDICT],type_ignores=[])),'real_supervisor_verdict','exec')

class VerdictTests(unittest.TestCase):
    def verdict(self,*,record=None,code=0,orphans=0,errors=None,worker='PASS'):
        r={'exit_code':0,'still_running':False,'crash_cleanup':{'status':'PASS','exit_code':0,'exit_observed_while_sta_alive':True}}
        if record:r.update(copy.deepcopy(record))
        env={'records':[r],'ownership_errors':errors or [],'summary':{'owned_orphans':orphans},
             'code':code,'normal':True,'data':{'status':worker}}
        exec(PROGRAM,env)
        return env['data']
    def test_verified_graceful_exit_is_pass(self):self.assertEqual(self.verdict()['status'],'PASS')
    def test_helper_failure_stays_fail(self):
        d=self.verdict(record={'crash_cleanup':{'status':'FAIL','exit_code':0,'exit_observed_while_sta_alive':True}})
        self.assertEqual(d['status'],'FAIL');self.assertEqual(d['error'],'OFFICE_OWNED_UNBOUND_CLEANUP_FAILED')
    def test_abnormal_and_missing_helper_exit_code_are_fail(self):
        for code in (1,259,None):
            with self.subTest(code=code):self.assertEqual(self.verdict(record={'crash_cleanup':{'status':'PASS','exit_code':code,'exit_observed_while_sta_alive':True}})['status'],'FAIL')
    def test_live_sta_proof_required(self):
        for proof in (False,None):
            with self.subTest(proof=proof):self.assertEqual(self.verdict(record={'crash_cleanup':{'status':'PASS','exit_code':0,'exit_observed_while_sta_alive':proof}})['status'],'FAIL')
    def test_emergency_termination_is_never_pass_even_with_zero_exit(self):
        d=self.verdict(record={'forced_termination_attempted':True})
        self.assertEqual(d['status'],'FAIL');self.assertEqual(d['error'],'OFFICE_GRACEFUL_EXIT_TIMEOUT')
    def test_actual_abnormal_process_exit_is_fail(self):self.assertEqual(self.verdict(record={'exit_code':1})['status'],'FAIL')
    def test_worker_failure_is_preserved(self):
        self.assertEqual(self.verdict(worker='FAIL')['status'],'FAIL')
        self.assertEqual(self.verdict(code=1)['status'],'FAIL')
    def test_unproven_ownership_or_living_orphan_is_fail(self):
        self.assertEqual(self.verdict(errors=[{'error':'NOT_OWNED'}])['status'],'FAIL')
        self.assertEqual(self.verdict(orphans=1)['status'],'FAIL')
    def test_invalid_seal_never_opens_or_quits_any_process(self):
        with tempfile.TemporaryDirectory() as folder:
            path=Path(folder)/'invalid.json'
            path.write_text(json.dumps({'record':{},'parent_pid':0,'parent_creation_time':'x','seal':'invalid'}),encoding='utf8')
            with patch('win32api.OpenProcess') as opened,patch('stage9_app.ownership.quit_owned_window') as quit_:
                with self.assertRaisesRegex(RuntimeError,'CLEANUP_PLAN_OR_PARENT_CHANGED'):cleanup_entry(path)
                opened.assert_not_called();quit_.assert_not_called()

if __name__=='__main__':unittest.main(verbosity=2)
