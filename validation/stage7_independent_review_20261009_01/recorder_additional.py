import sys
sys.dont_write_bytecode=True
import ast,json
from pathlib import Path
from copy import deepcopy
ROOT=Path(__file__).resolve().parents[2];OUT=Path(__file__).resolve().parent
sys.path.insert(0,str(ROOT))
from stage7_quality.core import verify_manifest
from stage7_quality.provenance import Stage7Error,digest,read,sha,Recorder
from stage7_quality import sources as S
M=read(ROOT/'validation/stage7_resume_20261009_175000/ASSEMBLY_01/STAGE7_SELECTED_TEST_MANIFEST.json')
FIX=OUT/'negative_fixtures'
results=read(OUT/'INDEPENDENT_TARGETED_TESTS.json')['tests']
# Execute only the three additional independent tests defined in the audit script.
# The earlier 22 tests and their exclusive fixtures are deliberately left intact.
tree=ast.parse((OUT/'negative_review.py').read_text(encoding='utf8'))
nodes=[n for n in tree.body if isinstance(n,ast.FunctionDef) and n.name in ('save','creates_output','additional')]
exec(compile(ast.Module(body=nodes,type_ignores=[]),str(OUT/'negative_review.py'),'exec'))
additional()
print(json.dumps({'total':len(results),'passed':sum(r['result']=='PASS' for r in results),'additional':results[-3:]},ensure_ascii=False))
