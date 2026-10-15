"""Fail-closed evidence finalizer. Review manifest is always the last output."""
import json
from pathlib import Path
import sys
sys.stdout.reconfigure(encoding='utf-8')
OUT=Path(__file__).resolve().parent
ROOT=OUT.parents[1]
sys.path.insert(0,str(ROOT))
from profiles_stage4.offline import read_json,sha
from profiles_stage4.storage import load_catalog,verify_workbook,DEFAULT_WORKBOOK,workbook_payload,dump,STATE,CATALOG,COMMIT


def validate():
    tests=read_json(OUT/'tests.json')
    before=read_json(OUT/'product_before.json')
    after=read_json(OUT/'product_after.json')
    product_paths=[p for p in (ROOT/'profiles_stage4').glob('*') if p.is_file()]+[
        ROOT/'stage4.py',ROOT/'Запустить Этап 4.cmd',ROOT/'STAGE4_README.md',OUT/'run_validation.py',OUT/'finalize.py']
    if {str(p.relative_to(ROOT)) for p in product_paths}!=set(before):
        raise ValueError('Product path set differs from exact tested files')
    current={p:sha(ROOT/p) for p in before}
    if before!=after or before!=current or tests['tested_product_hashes']!=before:
        raise ValueError('Current product does not match exact tested hashes')
    if tests['failed'] or not tests['protected_unchanged'] or tests['network_events'] or tests['production_events']:
        raise ValueError('Validation gate failed')
    passed={t['id'] for t in tests['tests'] if t['status']=='PASS'}
    if not set(range(1,44))<=passed:
        raise ValueError('Missing mandatory targeted tests')
    protected=read_json(OUT/'protected_before.json')
    if any(sha(p)!=h for p,h in protected.items()):
        raise ValueError('Protected Stage 1–3/source evidence changed')
    result=read_json(OUT/'build_result.json')
    if sha(DEFAULT_WORKBOOK)!=result['sha256']:
        raise ValueError('Workbook differs from tested output')
    cat=load_catalog()
    if cat.counts()!=result['counts']:
        raise ValueError('Current catalogue counts differ')
    reopen=verify_workbook(DEFAULT_WORKBOOK,workbook_payload(cat))
    return {'status':'PASS','tested_product_hashes':before,'product_before_equals_after_equals_current':True,
            'protected_count':len(protected),'protected_unchanged':True,'tests_passed':tests['passed'],
            'workbook':str(DEFAULT_WORKBOOK),'workbook_sha256':sha(DEFAULT_WORKBOOK),'reopen':reopen,
            'counts':cat.counts(),'network_requests':0,'production_access':0,
            'user_state_sha256':sha(DEFAULT_WORKBOOK.parent/STATE),
            'catalogue_sha256':sha(DEFAULT_WORKBOOK.parent/CATALOG),'commit_sha256':sha(DEFAULT_WORKBOOK.parent/COMMIT)}


if __name__=='__main__':
    try:
        result=validate()
        result['evidence_files']={str(p.relative_to(OUT)): {'sha256':sha(p),'size':p.stat().st_size}
            for p in sorted(OUT.rglob('*')) if p.is_file() and p.name!='STAGE4_REVIEW_MANIFEST.json'}
        dump(OUT/'STAGE4_REVIEW_MANIFEST.json',result)
        print(json.dumps({'status':'PASS','tests':result['tests_passed'],'workbook_sha256':result['workbook_sha256']},ensure_ascii=False))
    except (ValueError,OSError,KeyError) as exc:
        print('FINALIZER STOP: '+str(exc),file=sys.stderr)
        sys.exit(1)
