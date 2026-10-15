from audit import *
checks=[]
for rel in ('validation/stage9_092_20261012_01/EVIDENCE_MANIFEST.json','validation/work09_independent_20261012_01/EVIDENCE_MANIFEST.json'):
    p=ROOT/rel;m=read(p);expected={};changed=[]
    for field in ('files','external_files','product_bindings'):
        for name,value in m.get(field,{}).items():
            target=Path(name) if Path(name).is_absolute() else p.parent/name
            if target.name=='ВЗМК.xlsx':continue
            if field=='product_bindings' and (target==ROOT/'stage9.py' or target.parent==ROOT/'stage9_app'):
                changed.append({'path':str(target),'old_sha256':value,'current_sha256':sha(target)});continue
            expected[str(target)]=value
    actual=hashes([Path(k) for k in expected]);bad=[{'path':k,'expected':v,'actual':actual[k]} for k,v in expected.items() if actual[k]!=v]
    checks.append({'manifest':str(p),'manifest_sha256':sha(p),'references':len(expected),'mismatches':bad,'authorized_stage9_changes':changed})
save(OUT/'HANDOFF_READBACK.json',{'status':'PASS' if all(not c['mismatches'] for c in checks) else 'FAIL','checks':checks,'independent_runtime_proof':False})
assert all(not c['mismatches'] for c in checks),checks
save(OUT/'BASELINE_READY.json',{'status':'PASS','files':len(read(OUT/'PROTECTED_BEFORE.json')['files']),'exe_sha256':sha(EXE),'utc':datetime.datetime.now(datetime.timezone.utc).isoformat()})
print('Readback PASS',[(c['references'],len(c['authorized_stage9_changes'])) for c in checks],flush=True)
