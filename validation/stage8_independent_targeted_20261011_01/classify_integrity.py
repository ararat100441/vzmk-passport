"""Separate a live index-worker heartbeat from protected product/data evidence."""
import hashlib,json
from pathlib import Path
RUN=Path(__file__).resolve().parent;ROOT=RUN.parents[1]
def read(p):return json.loads(p.read_text(encoding='utf-8-sig'))
def save(p,v):p.write_text(json.dumps(v,ensure_ascii=False,indent=2),encoding='utf8')
def sha(p):return hashlib.sha256(p.read_bytes()).hexdigest()
raw=read(RUN/'SOURCE_INTEGRITY.json');save(RUN/'ALL_ACCESSIBLE_INTEGRITY.json',raw)
before=read(RUN/'PROTECTED_BEFORE.json');after=read(RUN/'PROTECTED_AFTER.json')
changed=raw['changed_or_unreadable'];assert len(changed)==1
excluded=[]
for p,error in changed.items():
    q=Path(p)
    assert q.is_relative_to(ROOT/'.codebase-index/index') and q.parent.name.startswith('background-worker.') and q.parent.name.endswith('.lease') and q.name.startswith('heartbeat.') and q.suffix=='.json'
    assert error=='SHA_MISMATCH'
    excluded.append({'path':p,'reason':'Live codebase-index worker lease heartbeat; volatile tool metadata, not product/source/business evidence','before':before[p],'after':after[p]})
scope_before={p:h for p,h in before.items() if p not in changed};scope_after={p:h for p,h in after.items() if p not in changed}
assert scope_before==scope_after and len(scope_before)==22627
save(RUN/'PROTECTED_DATA_BEFORE.json',scope_before);save(RUN/'PROTECTED_DATA_AFTER.json',scope_after)
save(RUN/'SOURCE_INTEGRITY.json',{'status':'PASS','protected_files':len(scope_before),'changed_or_unreadable':{},'current_product_matches':raw['current_product_matches'],'all_accessible_files':len(before),'excluded_volatile_tool_metadata':excluded,'all_accessible_raw_result':'ALL_ACCESSIBLE_INTEGRITY.json','before_manifest_sha256':sha(RUN/'PROTECTED_DATA_BEFORE.json'),'after_manifest_sha256':sha(RUN/'PROTECTED_DATA_AFTER.json')})
print('PROTECTED_DATA_PASS',len(scope_before),'VOLATILE_HEARTBEAT',len(excluded))
