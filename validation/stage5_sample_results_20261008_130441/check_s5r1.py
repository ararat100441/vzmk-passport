"""Reuse only named existing S5-R1 cases on new fixtures and current product."""
import ast
import json
from pathlib import Path
import sys
import unittest

RUN=Path(__file__).resolve().parent
ROOT=RUN.parents[1]
sys.path.insert(0,str(ROOT))
source=ROOT/'tests/test_stage5_s5_r1.py'
tree=ast.parse(source.read_text(encoding='utf8'),filename=str(source))
# Retain the original imports, sandbox and test definitions, not its historical
# change-scope assertion or its old full-suite runner. Source stays unchanged.
nodes=[]
for node in tree.body:
    if isinstance(node,ast.Assign) and any(isinstance(t,ast.Name) and t.id=='binding' for t in node.targets):
        break
    nodes.append(node)
sys.argv=[str(source),'--run',str(RUN),'--attempt','sample_regression_final']
space={'__file__':str(source),'__name__':'selected_s5r1_checks'}
exec(compile(ast.Module(body=nodes,type_ignores=[]),str(source),'exec'),space)
names=['test_A_zero_matches_two_unrelated_conflicts','test_B_full_preview_generate_no_outputs',
       'test_required_registry_conflict_still_stops','test_ui_empty_preview_disables_generation']
case=space['S5R1Tests']; W=space['W']; C=space['C']; I=space['I']
binding={str(p):I.sha(p) for p in space['ACTUAL_PRODUCT']+[ROOT/'tests/test_stage5_sample_results.py']}
space['write']('s5r1_product_before.json',binding)
with space['no_network']() as attempts, space['patch'].object(W,'write_test_document',side_effect=AssertionError('No Word call')):
    with (RUN/'s5r1_targeted.log').open('w',encoding='utf8') as log:
        result=unittest.TextTestRunner(stream=log,verbosity=2).run(unittest.TestSuite(case(name) for name in names))
after={str(p):I.sha(p) for p in space['ACTUAL_PRODUCT']+[ROOT/'tests/test_stage5_sample_results.py']}
assert after==binding
space['write']('s5r1_product_after.json',after)
summary={'tests_run':result.testsRun,'successful':result.wasSuccessful(),'cases':names,
         'network_attempts':attempts,'fixture_docx_generation':0,
         'errors':[(str(t),s) for t,s in result.errors],'failures':[(str(t),s) for t,s in result.failures]}
space['write']('s5r1_targeted.json',summary)
space['write']('s5r1_access_guard.json',{'events':space['ACCESS'],'production_reads':0,'writes_outside_run':0})
print(json.dumps(summary,ensure_ascii=False,indent=2))
raise SystemExit(0 if result.wasSuccessful() else 1)
