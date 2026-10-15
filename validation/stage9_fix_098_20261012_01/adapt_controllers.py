import json,hashlib
from pathlib import Path
OUT=Path(__file__).resolve().parent;ROOT=OUT.parents[1];OLD=ROOT/'validation/work09_targeted_097_20261012_01'
rows=[]
for name in ('common','e2e','negative','lifecycle','normal_overlap','overlap_crash_cleanup','window_close_guard','extended_guard','foreign_pumped','ownership_stop','viewer_snapshot','semantic_audit','receipts'):
    src=OLD/(name+'.py');s=src.read_text(encoding='utf8')
    s=s.replace('0.9.7','0.9.9').replace('stage9_dist_097','stage9_dist_099')
    if name=='common':s=s.replace("EXPECTED='cd140f337612c462903ecc35bb55173ba10839bb16fca61492c7ea7b65663cd7'","EXPECTED=read(OUT/'NEW_EXE_BINDING.json')['exe_sha256']")
    target=OUT/(name+'.py');target.write_text(s,encoding='utf8');rows.append({'original':str(src),'sha256':hashlib.sha256(src.read_bytes()).hexdigest(),'adapted':str(target)})
(OUT/'CONTROLLER_PROVENANCE.json').write_text(json.dumps({'kind':'Codex self-check controllers adapted from sealed Work scenarios; not independent Work PASS','controllers':rows},ensure_ascii=False,indent=2),encoding='utf8')
print('Controllers prepared; not executed')
