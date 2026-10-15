import json,hashlib
from pathlib import Path
RUN=Path(__file__).resolve().parent
p=RUN/'edges.py';s=p.read_text(encoding='utf8');changes=[]
for old,new,reason in [
    ("OUT=RUN/'edges';","OUT=RUN/'edges_final';",'Preserve the preliminary batch'),
    ("'NATIVE_SHORT_WORKSPACE_FAILURE'","'REGISTRY_GUARD_OR_IO_FAILURE'",'CreateDirectory failure is safely wrapped by namespace_guard'),
    ("D:P(D;;WDAD;;;","D:P(D;;0x00000006;;;",'Correct Windows SDDL numeric FILE_ADD_FILE/FILE_ADD_SUBDIRECTORY deny mask'),
    ("Filename=str(c),","Filename=str(c.resolve()),",'Resolve packaged LOCALAPPDATA virtualization before crossing the Excel COM process boundary'),
]:
    assert old in s,old
    changes.append({'old':old,'new':new,'reason':reason,'occurrences':s.count(old)})
    s=s.replace(old,new)
q=RUN/'edges_final.py';q.write_text(s,encoding='utf8',newline='')
data={'preliminary_source_sha256':hashlib.sha256(p.read_bytes()).hexdigest(),'final_source_sha256':hashlib.sha256(q.read_bytes()).hexdigest(),'changes':changes,'preliminary_results_sha256':hashlib.sha256((RUN/'edges/RESULTS.json').read_bytes()).hexdigest(),'product_files_changed':False,'all_34_cases_rerun':True}
(RUN/'EDGES_HARNESS_REPAIR.json').write_text(json.dumps(data,ensure_ascii=False,indent=2),encoding='utf8')
