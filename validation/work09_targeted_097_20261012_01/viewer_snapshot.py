from common import *
import pythoncom,win32api,win32process,gc
from stage9_app.ownership import native_application,process_image
launch=read(OUT/'GUI_LAUNCH.json');base=Path(SETUP['base'])/'VZMKPassport/diagnostics'
label=sys.argv[1];passport=sys.argv[2]
manifests=[read(p) for p in Path(launch['workspace']).rglob('MANIFEST.json')]
m=max((m for m in manifests if m['passport']==passport),key=lambda m: max(Path(r['path']).stat().st_mtime for r in m['records']))
expected={str(Path(r['path']).resolve()).casefold():r['sha256'] for r in m['records']}
starts=[]
for p in base.glob('viewer_*_START.json'):
    r=read(p)
    if r['launcher_pid']==launch['pid']:starts.append((p,r))
latest={}
for p,r in starts:
    if r['progid'] not in latest or r['creation_time']>latest[r['progid']][1]['creation_time']:latest[r['progid']]=(p,r)
assert len(latest)==2
snap=[];pythoncom.CoInitialize()
try:
    for p,r in latest.values():
        h=win32api.OpenProcess(0x100400,False,r['pid'])
        try:
            assert win32process.GetProcessTimes(h)['CreationTime'].isoformat()==r['creation_time'] and process_image(h)==r['image']
            assert r['actual_child'] and r['owned'] and r['handle_source']=='CREATE_PROCESS'
            items=[]
            with native_application(r,h) as app:
                col=app.Documents if r['progid']=='Word.Application' else app.Workbooks
                assert col.Count==3,(r,col.Count)
                for i in range(1,col.Count+1):
                    item=col(i);path=str(Path(item.FullName).resolve());assert item.ReadOnly and path.casefold() in expected,(path,item.ReadOnly)
                    assert sha(path)==expected[path.casefold()]
                    items.append({'path':path,'readonly':bool(item.ReadOnly),'sha256':sha(path)});item=None
                col=None
            app=None;gc.collect()
            snap.append({'identity':r,'START_evidence':str(p),'opened_physical_files':items})
        finally:h.Close()
finally:app=col=None;gc.collect();pythoncom.CoUninitialize()
assert sum(len(r['opened_physical_files']) for r in snap)==6
save(OUT/('GUI_VIEWER_'+label+'_SNAPSHOT.json'),{'status':'PASS','read_only_native_inspection':True,'passport':passport,'instances':snap,'physical_files':6,'logical_records':len(m['records']),'no_write_close_or_quit_by_controller':True})
print('GUI viewer',label,'3 Word / 3 Excel, ReadOnly, SHA unchanged PASS',flush=True)
