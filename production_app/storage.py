import contextlib,hashlib,json,os,stat
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
    return Path(os.environ.get('LOCALAPPDATA',str(Path.home())))/'VZMKPassportProduction'/'1.0.0'

def windows_file_identity(path):
    """Read authoritative metadata from one no-follow, zero-access handle.

    A Windows stat fallback may omit the link count. Zero is unknown, not
    evidence of a hardlink. Never infer safety when this native query fails.
    """
    import win32file,win32con
    try:
        handle=win32file.CreateFile(str(path),0,
            win32con.FILE_SHARE_READ|win32con.FILE_SHARE_WRITE|win32con.FILE_SHARE_DELETE,
            None,win32con.OPEN_EXISTING,0x00200000,None)  # OPEN_REPARSE_POINT
        try:
            info=win32file.GetFileInformationByHandle(handle)
            return {'attributes':info[0],'links':info[7],
                    'volume':info[4],'file_id':(info[8]<<32)|info[9]}
        finally:handle.Close()
    except OSError as exc:
        raise Stop('Не удалось проверить идентичность файла Windows: '+str(path)+
                   '. Закройте редактирование файла и повторите. '+str(exc)) from exc

def safe(path):
    path=Path(path).absolute()
    for p in (path,*path.parents):
        try:metadata=p.lstat()
        except FileNotFoundError:continue
        if stat.S_ISLNK(metadata.st_mode) or getattr(metadata,'st_file_attributes',0)&0x400:
            raise Stop('Ссылка/junction/reparse недопустима: '+str(p))
    if path.is_file():
        if os.name=='nt':
            identity=windows_file_identity(path)
            if identity['attributes']&0x400:raise Stop('Ссылка/reparse недопустима: '+str(path))
            links=identity['links']
        else:links=path.stat().st_nlink
        if links>1:raise Stop('Hardlink недопустим: '+str(path))
        if links!=1:raise Stop('Не удалось подтвердить единственное имя файла: '+str(path))
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
