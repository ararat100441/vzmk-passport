import hashlib,json,sys,io
from pathlib import Path
from zipfile import ZipFile
import xml.etree.ElementTree as ET
RUN=Path(__file__).resolve().parent
def read(p):return json.loads(Path(p).read_text(encoding='utf-8-sig'))
def save(p,x):Path(p).write_text(json.dumps(x,ensure_ascii=False,indent=2),encoding='utf8')

if len(sys.argv)>1 and sys.argv[1]=='rot':
    import pythoncom,win32com.client,win32process
    pythoncom.CoInitialize();rot=pythoncom.GetRunningObjectTable();ctx=pythoncom.CreateBindCtx(0);matches=[]
    crash=read(RUN/'cases/actual_process_death_before_save.json')['detail']['sandbox']
    for m in rot.EnumRunning():
        try:name=m.GetDisplayName(ctx,None)
        except Exception:continue
        if str(RUN).casefold() not in name.casefold():continue
        item={'moniker':name}
        try:
            obj=win32com.client.Dispatch(rot.GetObject(m).QueryInterface(pythoncom.IID_IDispatch))
            app=obj.Application;books=[b.FullName for b in app.Workbooks]
            item.update(pid=win32process.GetWindowThreadProcessId(app.Hwnd)[1],visible=bool(app.Visible),books=books)
            # Only the exact own crash fixture, with no other workbooks and no visible UI, may be closed.
            own=bool(books) and all(Path(b).is_relative_to(Path(crash)) for b in books) and not app.Visible
            if own and '--cleanup' in sys.argv:
                for b in list(app.Workbooks):b.Close(SaveChanges=False)
                app.Quit();item['cleanup']='OWN_CRASH_EXCEL_ONLY_CLOSED_WITHOUT_SAVE'
        except Exception as exc:item['error']=str(exc)
        matches.append(item)
    save(RUN/('EXCEL_ROT_CLEANUP.json' if '--cleanup' in sys.argv else 'EXCEL_ROT_BEFORE_CLEANUP.json'),matches)
    print(json.dumps(matches,ensure_ascii=False));sys.exit()

data=read(RUN/'POSITIVE_NATIVE.json');before=data['backup']['path'];after=Path(data['sandbox'])/'registry/УЗК_TEST.xlsm'
def short(tag):return tag.rsplit('}',1)[-1]
def diff(a,b,path='/'):
    out=[]
    if a.tag!=b.tag:return [{'path':path,'tag_before':a.tag,'tag_after':b.tag}]
    for k in set(a.attrib)|set(b.attrib):
        if a.get(k)!=b.get(k):out.append({'path':path+'/@'+short(k),'before':a.get(k),'after':b.get(k)})
    if (a.text or '')!=(b.text or ''):out.append({'path':path+'/text()','before':a.text,'after':b.text})
    def children(parent):
        result={};counts={}
        for x in parent:
            name=x.tag;counts[name]=counts.get(name,0)+1
            identity=next((k+'='+x.get(k) for k in ('r','name','Id','id') if x.get(k)),str(counts[name]))
            result[name+'['+identity+']']=x
        return result
    ac=children(a);bc=children(b)
    for k in sorted(set(ac)|set(bc)):
        x=ac.get(k);y=bc.get(k);suffix=short(k)
        if x is None or y is None:out.append({'path':path+'/'+suffix,'removed':ET.tostring(x,encoding='unicode') if x is not None else None,'added':ET.tostring(y,encoding='unicode') if y is not None else None})
        else:out.extend(diff(x,y,path+'/'+suffix))
    return out
with ZipFile(before) as a,ZipFile(after) as b:
    parts=[]
    for n in sorted(set(a.namelist())|set(b.namelist())):
        x=a.read(n) if n in a.namelist() else None;y=b.read(n) if n in b.namelist() else None
        if x==y:continue
        d={'part':n,'before_sha256':hashlib.sha256(x).hexdigest() if x else None,'after_sha256':hashlib.sha256(y).hexdigest() if y else None}
        if n.endswith(('.xml','.rels')) and x and y:d['xml_differences']=diff(ET.fromstring(x),ET.fromstring(y),short(ET.fromstring(x).tag))
        parts.append(d)
    result={'backup':before,'closed_registry':str(after),'part_count_before':len(a.namelist()),'part_count_after':len(b.namelist()),'raw_changed_parts':parts,
            'styles_xml_byte_identical':a.read('xl/styles.xml')==b.read('xl/styles.xml'),
            'workbook_relationships_byte_identical':a.read('xl/_rels/workbook.xml.rels')==b.read('xl/_rels/workbook.xml.rels')}
save(RUN/'OOXML_RAW_DIFF.json',result);print(json.dumps(result,ensure_ascii=False))
