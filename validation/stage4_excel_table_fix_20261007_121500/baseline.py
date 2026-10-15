from pathlib import Path
import hashlib,json,shutil,sys
sys.stdout.reconfigure(encoding='utf-8')
OUT=Path(__file__).resolve().parent; ROOT=OUT.parents[1]
sys.path.insert(0,str(OUT))
from xlsx_tables import inspect
def sha(path): return hashlib.sha256(Path(path).read_bytes()).hexdigest()
def save(name,value): (OUT/name).write_text(json.dumps(value,ensure_ascii=False,indent=2),encoding='utf-8')
if __name__=='__main__':
    folder=ROOT/'DEV/Справочники'
    expected={'Профили.xlsx':'2a07eeacd56793df749ddb848563336a413f54beaae9b6d002e438c35452b98e',
        'profiles.user.json':'a6031a6f08502f4cd1e6f99ef0ac27d284f2f2a3e5d1578a006fb134448dc67b',
        'profiles.catalog.json':'350564a44f4c0577ed540cbcf595f397ca2b33ab5e7b7bf2dc4e629d8021aa6d',
        'profiles.build.json':'fb075fd8a622f72d48d231ae9391baba7bcd3d37f297abff38bb3a0da2959508'}
    current={str(folder/n):sha(folder/n) for n in expected}
    assert current=={str(folder/n):h for n,h in expected.items()},'STOP: real DEV baseline mismatch'
    assert sha(ROOT/'DEV/ВЗМК_DEV.xlsx')=='8e56f68518c70cade9efb09084dd67cd5d72e07e3fdc813b8515030e9cb98e80'
    previous=ROOT/'validation/stage4_fix_m1_m2_20261007_090000'
    manifest=json.loads((previous/'REVIEW_MANIFEST.json').read_text(encoding='utf-8'))
    protected=manifest['protected_hashes'].copy()
    mutable={str(ROOT/'profiles_stage4/storage.py'),str(ROOT/'profiles_stage4/workbook.mjs')}
    for relative,digest in manifest['product_hashes'].items():
        path=ROOT/relative
        assert sha(path)==digest,relative
        if str(path) not in mutable: protected[str(path)]=digest
    for name in ['stage4_fix_m1_m2_20261007_090000','stage4_targeted_independent_rereview_20261007_112239',
                 'stage4_targeted_rereview_20261007_current_2a07ee']:
        for path in (ROOT/'validation'/name).rglob('*'):
            if path.is_file(): protected[str(path)]=sha(path)
    assert all(sha(p)==h for p,h in protected.items()),'Protected evidence changed'
    save('generation_before.json',current); save('protected_before.json',protected)
    snapshot=OUT/'baseline_generation'; snapshot.mkdir()
    for name in expected: shutil.copy2(folder/name,snapshot/name)
    forensic=inspect(folder/'Профили.xlsx',OUT/'pre_fix_xml')
    save('pre_fix_structure.json',forensic)
    logs=sorted((Path.home()/'AppData/Local/Temp').glob('error*.xml'))
    saved=[]
    for log in logs:
        try:
            data=log.read_bytes()
            if 'table1.xml' in data.decode('utf-8-sig',errors='replace'):
                destination=OUT/'repair_logs'/log.name; destination.parent.mkdir(exist_ok=True)
                destination.write_bytes(data); saved.append({'source':str(log),'copy':str(destination),'sha256':sha(destination)})
        except OSError: pass
    save('existing_repair_logs.json',{'found':saved,'searched':str(Path.home()/'AppData/Local/Temp/error*.xml')})
    print(json.dumps({'baseline':'MATCH','recovered_copy_saved':False,'protected_files':len(protected),
        'forensic_issues':forensic['issues'],'existing_repair_logs':len(saved)},ensure_ascii=False))
