import sys,json,shutil
from pathlib import Path
RUN=Path(__file__).resolve().parent;ROOT=RUN.parents[1];sys.path.insert(0,str(ROOT))
from stage8_closure import native,office,storage as S
p=next(RUN.glob('sandboxes/*/transactions/*/restore_*/verified.xlsm'))
short=RUN/'SAME_BYTES_SHORT.xlsm';shutil.copyfile(p,short)
out={'long_path':str(p),'long_length':len(str(p)),'short_path':str(short),'short_length':len(str(short)),'sha_equal':S.sha(p)==S.sha(short),'sha256':S.sha(p)}
with office.diagnostic_scope(RUN/'path_probe_office'):
    for name,path in [('long',p),('short',short)]:
        try:out[name]={'status':'PASS','native':native.inspect(path)}
        except Exception as e:out[name]={'status':'STOP','error':office._error(e)}
(RUN/'PATH_LENGTH_PROOF.json').write_bytes(json.dumps(out,ensure_ascii=False,indent=2).encode('utf8'))
print(json.dumps(out,ensure_ascii=False))
