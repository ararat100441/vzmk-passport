from common import *
from stage9_app.ownership import owned_office
import traceback
p=OUT/'UNSUPERVISED_EMPTY_GUARD_DIAGNOSTIC.json';error=None
try:
    with owned_office(p,supervised=False):pass
except Exception as exc:error={'type':type(exc).__name__,'error':str(exc),'traceback':traceback.format_exc()}
assert error and error['type']=='FileExistsError' and read(p)['owned_orphans']==0
save(OUT/'UNSUPERVISED_EMPTY_GUARD_REPRO.json',{'status':'REPRODUCED','source_only':True,'office_launches':0,'expected_normal_exit':True,'actual_exception':error,'initial_diagnostic':str(p),'initial_diagnostic_sha256':sha(p),'source_sha256':sha(ROOT/'stage9_app/ownership.py')})
print('Deterministic unsupervised normal teardown double-write reproduced; Office launches 0',flush=True)
