import hashlib, json, sys, time
from pathlib import Path

RUN = Path(__file__).resolve().parent
ROOT = RUN.parents[1]
PACKAGE = ROOT/'validation/stage8_20261010_03'
def load(p): return json.loads(Path(p).read_text(encoding='utf-8-sig'))
def sha(p): return hashlib.sha256(Path(p).read_bytes()).hexdigest()
def save(p, data): Path(p).write_text(json.dumps(data,ensure_ascii=False,indent=2),encoding='utf8')

phase=sys.argv[1]
baseline=load(ROOT/'validation/stage8_20261010_01/SOURCE_BEFORE.json')
reported_after=load(PACKAGE/'SOURCE_AFTER.json')
manifest=load(PACKAGE/'REVIEW_MANIFEST.json')
groups={k:v for k,v in manifest.items() if isinstance(v,dict) and all(isinstance(h,str) and len(h)==64 for h in v.values())}
paths=set(baseline)|set(reported_after)
for values in groups.values(): paths.update(values)
paths.update(str(p.resolve()) for p in (ROOT/'stage8_closure').glob('*') if p.is_file())
paths.add(str(PACKAGE/'REVIEW_MANIFEST.json'))
current={}; errors=[]
start=time.monotonic()
for p in sorted(paths):
    try: current[p]=sha(p)
    except Exception as exc: errors.append({'path':p,'error':repr(exc)})
def changes(expected):
    return [{'path':p,'expected':h,'actual':current.get(p)} for p,h in expected.items() if current.get(p)!=h]
result={'phase':phase,'seconds':round(time.monotonic()-start,2),'protected_count':len(baseline),
 'protected_vs_checkpoint':changes(baseline),'reported_after_vs_checkpoint':[p for p in set(baseline)|set(reported_after) if baseline.get(p)!=reported_after.get(p)],
 'manifest_groups':{k:{'count':len(v),'mismatches':changes(v)} for k,v in groups.items()},'read_errors':errors,
 'report_manifest_sha256':sha(PACKAGE/'REVIEW_MANIFEST.json')}
integrity=load(PACKAGE/'SOURCE_INTEGRITY.json')
result['source_table_sha_checks']={name:sha(integrity[name+'_path'])==integrity[name+'_sha256'] for name in ('before','after')}
if phase=='AFTER': result['review_changes']=changes(load(RUN/'HASH_BEFORE.json'))
save(RUN/('HASH_'+phase+'.json'),current)
save(RUN/('INTEGRITY_'+phase+'.json'),result)
print(json.dumps({k:v for k,v in result.items() if k!='manifest_groups'},ensure_ascii=False))
print(json.dumps(result['manifest_groups'],ensure_ascii=False))
