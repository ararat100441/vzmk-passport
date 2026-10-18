import contextlib,hashlib,json,os
from pathlib import Path
from uuid import uuid4

class Stop(RuntimeError): pass

def sha(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()

def digest(data):
    return hashlib.sha256(json.dumps(data,sort_keys=True,ensure_ascii=False,default=str).encode('utf8')).hexdigest()

def read(path):
    return json.loads(Path(path).read_text(encoding='utf8'))

def save(path,data):
    path=Path(path);path.parent.mkdir(parents=True,exist_ok=True)
    temp=path.with_name('.'+uuid4().hex+'.tmp')
    with temp.open('x',encoding='utf8') as f:
        json.dump(data,f,ensure_ascii=False,indent=2,default=str);f.flush();os.fsync(f.fileno())
    os.replace(temp,path)

def local_root():
    return Path(os.environ.get('LOCALAPPDATA',str(Path.home())))/'VZMKPassportProduction'

def safe(path):
    path=Path(path).absolute()
    for p in (path,*path.parents):
        if p.exists() and (p.is_symlink() or p.is_junction()):raise Stop('Ссылка/junction недопустима: '+str(p))
    if path.is_file() and path.stat().st_nlink!=1:raise Stop('Hardlink недопустим: '+str(path))
    return path.resolve()

def contained(path,root):
    path,root=safe(path),safe(root)
    if not path.is_relative_to(root):raise Stop('Путь вне папки объекта: '+str(path))
    return path

@contextlib.contextmanager
def mutex(registry):
    import win32event,win32api,winerror
    name='Local\\VZMKPassport_'+hashlib.sha256(str(safe(registry)).casefold().encode()).hexdigest()
    h=win32event.CreateMutex(None,False,name)
    acquired=win32event.WaitForSingleObject(h,0) in (0,128)
    try:
        if not acquired:raise Stop('Реестр занят другим экземпляром программы.')
        yield
    finally:
        if acquired:win32event.ReleaseMutex(h)
        win32api.CloseHandle(h)

def pins(paths):return {str(safe(p)):sha(p) for p in paths}

def verify_pins(values):
    for p,h in values.items():
        if sha(safe(p))!=h:raise Stop('SHA-256 изменился: '+p)
