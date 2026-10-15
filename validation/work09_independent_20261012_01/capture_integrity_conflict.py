import json,hashlib
from pathlib import Path
from datetime import datetime,timezone
OUT=Path(__file__).resolve().parent
pins=json.loads((OUT/'PROTECTED_BEFORE.json').read_text(encoding='utf8'))
target=next(Path(p) for p in pins if Path(p).name=='ВЗМК.xlsx' and Path(p).parent.name=='Работа')
def sha():return hashlib.sha256(target.read_bytes()).hexdigest()
stat1=target.stat();first=sha();second=sha();stat2=target.stat()
d={'schema':'WORK09_PROTECTED_FILE_CONFLICT_V1','timestamp_utc':datetime.now(timezone.utc).isoformat(),'path':str(target),'expected_sha256':pins[str(target)],'actual_sha256':second,'consecutive_reads_equal':first==second,'size_bytes':stat2.st_size,'mtime_ns':stat2.st_mtime_ns,'stat_stable_during_read':(stat1.st_size,stat1.st_mtime_ns)==(stat2.st_size,stat2.st_mtime_ns),'writer':'UNKNOWN','review_did_not_write_restore_or_rebaseline_this_file':True,'previous_successful_snapshot':'INDEPENDENT_AFTER.json','source_integrity_gate':'FAIL','product_defect_attribution':False}
with (OUT/'FINAL_INTEGRITY_CONFLICT.json').open('x',encoding='utf8') as f:json.dump(d,f,ensure_ascii=False,indent=2)
print(json.dumps(d,ensure_ascii=False,indent=2),flush=True)
