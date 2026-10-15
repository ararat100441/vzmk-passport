"""Reuse only the prepared artificial integration protocol with exact new EXE."""
from common import *
WORKSPACE=Path(read(OUT/'WORKSPACE.json')['workspace'])
source=PRIOR/'exe_stage8.py';protocol=source.read_text(encoding='utf8')
old="d,p,s=task('EXE_STAGE8',name,expected,**args)"
assert protocol.count(old)==1
protocol=protocol.replace(old,"d,p=task('S8_'+name,name,expected,workspace=WORKSPACE,**args)")
copy=OUT/'stage8_adapted_protocol.py'
with copy.open('x',encoding='utf8') as f:f.write(protocol)
save(OUT/'STAGE8_CONTROLLER_PROVENANCE.json',{'source':str(source),'source_sha256':sha(source),'copy':str(copy),'copy_sha256':sha(copy),'change':'controller task signature only; current workspace and exact0.9.10 EXE. No prior EXE launched.'})
exec(compile(protocol,str(copy),'exec'),globals())
