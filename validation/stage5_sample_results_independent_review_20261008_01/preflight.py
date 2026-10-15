import ast
import difflib
import hashlib
import json
from pathlib import Path

OUT = Path(__file__).resolve().parent
ROOT = OUT.parent.parent
IMPL = ROOT / 'validation/stage5_sample_results_20261008_130441'
FORBIDDEN = {'взмк.xlsx', 'узк.xlsm'}

def sha(path):
    path=Path(path)
    assert path.name.casefold() not in FORBIDDEN
    return hashlib.sha256(path.read_bytes()).hexdigest()

def read(path):
    return json.loads(Path(path).read_text(encoding='utf-8-sig'))

def write(name, value):
    (OUT/name).write_text(json.dumps(value,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')

manifest=read(IMPL/'REVIEW_MANIFEST.json')
artifact=Path(manifest['output']).resolve()
assert artifact.is_file(), 'STOP: exact artifact missing'
assert sha(artifact)==manifest['output_sha256']=='8d3a2769baf6d381a0675804ef7f0806896f1b611e1498ea58677d3fdac1fd1b', 'STOP: exact artifact SHA mismatch'
current={p:sha(p) for p in manifest['product_current']}
assert current==manifest['product_current']==manifest['product_after']
assert current==manifest['test_binding_before']==manifest['test_binding_after'], 'STOP: tested bytes differ from current'
for name in ('product_before_tests.json','product_after_tests.json','product_current.json','s5r1_product_before.json','s5r1_product_after.json'):
    assert read(IMPL/name)==current, name
generation=read(IMPL/'product_before_generation.json')
assert all(current[p]==h for p,h in generation.items() if Path(p).suffix=='.py' and '/tests/' not in p.replace('\\','/'))

evidence={Path(e['path']).as_posix():e for e in manifest['evidence']}
actual={p.relative_to(IMPL).as_posix():p for p in IMPL.rglob('*') if p.is_file() and p.name!='REVIEW_MANIFEST.json'}
assert set(evidence)==set(actual)
for rel,item in evidence.items():
    assert sha(actual[rel])==item['sha256'], rel
    assert actual[rel].stat().st_size==item['size'], rel
assert (IMPL/'REVIEW_MANIFEST.json').stat().st_mtime_ns>=max(p.stat().st_mtime_ns for p in actual.values())

baseline=read(IMPL/'baseline.json')
allowed={str(ROOT/p) for p in ('uzk_stage5/word.py','uzk_stage5/gui.py','STAGE5_DEV.md')}
baseline_now={p:sha(p) for p in baseline['protected_before']}
changed={p for p,h in baseline['protected_before'].items() if baseline_now[p]!=h}
assert changed==allowed, changed
assert manifest['product_before']==baseline['product_before']
declared=read(IMPL/'changed_files.json')
assert {e['path'] for e in declared}==allowed|{str(ROOT/'uzk_stage5/result_samples.py'),str(ROOT/'tests/test_stage5_sample_results.py')}
for entry in declared:
    assert current[entry['path']]==entry['after']
    assert entry['before']==manifest['product_before'].get(entry['path'])
before_product=manifest['product_before']
assert {p for p in before_product if current[p]!=before_product[p]}==allowed
assert set(current)-set(before_product)=={str(ROOT/'uzk_stage5/result_samples.py'),str(ROOT/'tests/test_stage5_sample_results.py')}
assert sha(baseline['existing_test_docx']['path'])==baseline['existing_test_docx']['sha256']=='7723be279f6e776a472fb33f0b1bc153aecd14fc04e5b319fae73f89f2960661'

recovered={}
diffs=[]

def bound_bytes(text, expected):
    for candidate in (text, text.rstrip('\n'), text.rstrip('\n')+'\n'):
        for newline in ('\n','\r\n'):
            payload=candidate.replace('\n',newline).encode('utf-8')
            for prefix in (b'',b'\xef\xbb\xbf'):
                if hashlib.sha256(prefix+payload).hexdigest()==expected:
                    return prefix+payload
    raise AssertionError('Recovered source does not match accepted SHA')

for p in sorted(allowed):
    old_sha=before_product[p]
    candidates=[Path(q) for q,h in baseline['protected_before'].items() if h==old_sha and q!=p]
    new_text=Path(p).read_text(encoding='utf-8-sig')
    rel=Path(p).relative_to(ROOT).as_posix()
    if candidates:
        old=candidates[0]
        assert sha(old)==old_sha
        old_bytes=old.read_bytes()
        origin=str(old)
    elif rel=='uzk_stage5/word.py':
        patch=(IMPL/'TARGETED_CHANGE.diff').read_text(encoding='utf-8-sig').splitlines(True)
        chunks=[];chunk=[]
        for line in patch:
            if line.startswith('@@'):
                if chunk:chunks.append(chunk)
                chunk=[]
            elif line.startswith(('---','+++')):continue
            elif line[:1] in (' ','+','-'):chunk.append(line)
        if chunk:chunks.append(chunk)
        old_text=new_text
        for chunk in chunks:
            current_chunk=''.join(l[1:] for l in chunk if l[0] in (' ','+'))
            prior_chunk=''.join(l[1:] for l in chunk if l[0] in (' ','-'))
            assert old_text.count(current_chunk)==1
            old_text=old_text.replace(current_chunk,prior_chunk,1)
        old_bytes=bound_bytes(old_text,old_sha)
        origin='Independently reverse TARGETED_CHANGE.diff; recovered bytes bound to accepted SHA'
    elif rel=='uzk_stage5/gui.py':
        prior=ROOT/'validation/stage5_S5_R1_independent_rereview_20261008_01/recovered_previous_source/uzk_stage5/gui.py'
        old_banner=next(l for l in prior.read_text(encoding='utf-8-sig').splitlines(True) if 'ttk.Label(panel, text=' in l)
        new_banner=next(l for l in new_text.splitlines(True) if 'ttk.Label(panel, text=' in l)
        old_bytes=bound_bytes(new_text.replace(new_banner,old_banner,1),old_sha)
        origin='Accepted previous GUI banner restored; recovered bytes bound to accepted S5-R1 SHA'
    else:
        # The guide is documentation. Its prior SHA is authoritative; no prior text
        # snapshot exists. Review its complete current text against the requested
        # contract rather than fabricate a previous version.
        recovered[rel]={'before_sha256':old_sha,'current_sha256':sha(p),
                        'prior_text_snapshot_available':False,'assessment':'Complete current guide reviewed; changes confined to sample/test-only instructions and evidence handoff.'}
        continue
    old_text=old_bytes.decode('utf-8-sig').replace('\r\n','\n')
    target=OUT/'accepted_before'/rel
    target.parent.mkdir(parents=True,exist_ok=True)
    target.write_bytes(old_bytes)
    recovered[rel]={'accepted_snapshot':origin,'before_sha256':old_sha,'current_sha256':sha(p)}
    diffs.extend(difflib.unified_diff(old_text.splitlines(True),new_text.splitlines(True),fromfile='BEFORE/'+rel,tofile='CURRENT/'+rel))
    if rel.endswith('.py'):
        def defs(source):
            tree=ast.parse(source)
            return {n.name:ast.dump(n,include_attributes=False) for n in ast.walk(tree) if isinstance(n,ast.FunctionDef)}
        a,b=defs(old_text),defs(new_text)
        recovered[rel]['changed_functions']=[k for k in a if a[k]!=b.get(k)]
        recovered[rel]['added_functions']=sorted(set(b)-set(a))
for entry in declared:
    if entry['before'] is None:
        rel=Path(entry['path']).relative_to(ROOT).as_posix()
        diffs.extend(difflib.unified_diff([],Path(entry['path']).read_text(encoding='utf-8').splitlines(True),fromfile='BEFORE/'+rel,tofile='CURRENT/'+rel))
(OUT/'ACTUAL_PRODUCT.diff').write_text(''.join(diffs),encoding='utf-8')
assert recovered['uzk_stage5/word.py']['changed_functions']==['write_test_document']
assert recovered['uzk_stage5/gui.py']['changed_functions']==['__init__']

protected={**baseline_now,**current}
protected.update({str(p.resolve()):sha(p) for p in IMPL.rglob('*') if p.is_file()})
write('protected_before.json',protected)
write('product_before_review.json',current)
summary={'status':'PASS','current_files':len(current),'tested_BEFORE_AFTER_CURRENT_equal':True,
         'implementation_evidence_files':len(evidence),'manifest_written_last_verified':True,
         'implementation_manifest_sha256':sha(IMPL/'REVIEW_MANIFEST.json'),
         'baseline_protected_files':len(baseline_now),'allowed_changed':sorted(allowed),
         'added_files':sorted(set(current)-set(before_product)),
         'previous_protected_except_authorized_changes_unchanged':True,
         'protected_for_review_count':len(protected),'artifact':str(artifact),'artifact_sha256':sha(artifact),
         'old_test_docx':baseline['existing_test_docx'],'accepted_source_recovery':recovered}
write('preflight.json',summary)
print(json.dumps(summary,ensure_ascii=False,indent=2))
