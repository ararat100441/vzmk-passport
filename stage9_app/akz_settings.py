"""Lossless numeric identity and run-bound TEST settings provenance."""
from decimal import Decimal
from itertools import product
from pathlib import Path
import hashlib,json,math,re
from stage8_closure import storage as S
from .workspace import load

CANONICAL='AKZ_NUMERIC_CANONICAL_V1'
SNAPSHOT='STAGE9_FROZEN_AKZ_SETTINGS_V1'

def canonical_json(value):
    if value is None or type(value) is bool or isinstance(value,str):return json.dumps(value,ensure_ascii=False)
    if type(value) in (int,float):
        if not math.isfinite(value):raise S.Stop('AKZ_RANGE_NUMBER_INVALID')
        d=Decimal(str(value))
        if d==0:return '0'
        text=format(d,'f')
        return text.rstrip('0').rstrip('.') if '.' in text else text
    if isinstance(value,list):return '['+','.join(canonical_json(v) for v in value)+']'
    if isinstance(value,dict) and all(isinstance(k,str) for k in value):
        return '{'+','.join(json.dumps(k,ensure_ascii=False)+':'+canonical_json(value[k]) for k in sorted(value))+'}'
    raise S.Stop('AKZ_CANONICAL_TYPE_INVALID')

def normalize(value):return json.loads(canonical_json(value))
def digest(value):return hashlib.sha256(canonical_json(value).encode('utf8')).hexdigest()

def gui_number(text):
    try:d=Decimal(text.strip().replace(',','.'))
    except Exception as exc:raise S.Stop('AKZ_RANGE_NUMBER_INVALID') from exc
    if not d.is_finite():raise S.Stop('AKZ_RANGE_NUMBER_INVALID')
    if d==d.to_integral():return int(d)
    value=float(d)
    if not math.isfinite(value) or Decimal(str(value))!=d:raise S.Stop('AKZ_NUMBER_PRECISION_LOSS_STOP')
    return value

def legacy_climate_digests(cfg):
    """Six endpoints: at most 64 exact int/float representations, no rounding."""
    keys=('air_c','surface_c','relative_humidity');options=[]
    for key in keys:
        for value in cfg[key]:
            d=Decimal(str(value));options.append((int(d),float(d)) if d==d.to_integral() else (value,))
    result=set()
    for values in product(*options):
        candidate=dict(cfg)
        for i,key in enumerate(keys):candidate[key]=list(values[i*2:i*2+2])
        result.add(S.digest(candidate))
    return result

def run_path(root,run):
    root,_=load(root);run=S.no_links(run)
    if run.parent!=root/'validation' or not re.fullmatch(r's9_[0-9a-f]{12}',run.name) or not run.is_dir():raise S.Stop('AKZ_SNAPSHOT_RUN_INVALID')
    return root,run

def freeze(root,run,expected):
    from . import akz_auto as A
    root,run=run_path(root,run);live=A.settings_path(root)
    with S.frozen([live]):
        if S.sha(live)!=expected:raise S.Stop('AKZ_SETTINGS_CHANGED_AFTER_PREVIEW')
        raw=live.read_bytes();cfg=A.validate_settings(json.loads(raw),A.profiles(root))
        directory=S.no_links(run/'inputs');directory.mkdir(exist_ok=True);target=S.no_links(directory/'AKZ_SETTINGS.json')
        if target.exists():raise S.Stop('AKZ_FROZEN_SETTINGS_EXISTS_NO_OVERWRITE')
        S.exclusive_bytes(target,raw)
        import win32api,win32con
        win32api.SetFileAttributes(str(target),win32api.GetFileAttributes(str(target))|win32con.FILE_ATTRIBUTE_READONLY)
    return {'schema':SNAPSHOT,'path':str(target),'sha256':expected,'canonical_digest':digest(cfg)}

def verify_snapshot(root,run,ref):
    from . import akz_auto as A
    root,run=run_path(root,run)
    if not isinstance(ref,dict) or set(ref)!={'schema','path','sha256','canonical_digest'} or ref['schema']!=SNAPSHOT:raise S.Stop('AKZ_FROZEN_SETTINGS_REFERENCE_DAMAGED')
    path=S.no_links(ref['path'])
    if path!=run/'inputs/AKZ_SETTINGS.json' or not path.is_file() or S.sha(path)!=ref['sha256']:raise S.Stop('AKZ_FROZEN_SETTINGS_SHA_CONFLICT')
    cfg=A.validate_settings(S.read(path),A.profiles(root))
    if digest(cfg)!=ref['canonical_digest']:raise S.Stop('AKZ_FROZEN_SETTINGS_DIGEST_CONFLICT')
    return cfg

def legacy_evidence(path,expected):
    """Recover exact old settings bytes in memory; never rewrite old evidence."""
    path=S.no_links(path);root=path.parents[2];load(root)
    if path!=root/'preferences/akz_auto/settings.json':raise S.Stop('LEGACY_AKZ_SETTINGS_PATH_INVALID')
    if path.exists() and S.sha(path)==expected:
        return {'status':'LEGACY_SETTINGS_EXACT_SHA_VERIFIED','source':str(path),'sha256':expected}
    history=S.no_links(path.parent/'history')
    for candidate in sorted(history.glob('*.json')):
        record=S.read(S.no_links(candidate))
        if 'settings' not in record:continue
        text=json.dumps(record['settings'],ensure_ascii=False,indent=2,default=str)
        # workspace.save uses text mode: Windows writes CRLF; LF evidence also exists.
        # Accept either writer representation only when its exact bytes match the pinned SHA.
        for ending,newline in (('LF','\n'),('CRLF','\r\n')):
            raw=text.replace('\n',newline).encode('utf8')
            if hashlib.sha256(raw).hexdigest()==expected:
                return {'status':'LEGACY_SETTINGS_RECOVERED_EXACT_SHA','source':str(candidate),'evidence_sha256':S.sha(candidate),'sha256':expected,'old_files_rewritten':False,'recovered_line_endings':ending}
    raise S.Stop('LEGACY_AKZ_SETTINGS_UNPROVEN: исходные настройки не подтверждены evidence. Создайте новый TEST-комплект.')
