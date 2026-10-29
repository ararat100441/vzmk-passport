"""Standalone production GUI and isolated native Office workers."""
import argparse,json,os,sys,traceback
from pathlib import Path
sys.dont_write_bytecode=True

def task(name,request):
    from production_app import core,model,office,VERSION
    if name=='smoke':
        import tkinter as tk
        root=tk.Tk();root.withdraw();root.update();root.destroy()
        return {'status':'PASS','version':VERSION,'frozen':bool(getattr(sys,'frozen',False)),'gui':True}
    if name=='forms':return {'status':'PASS','folder':office.create_forms(request['folder'])}
    if name=='catalog':return {'status':'PASS','path':model.draft_catalog(request['vom'],request['target'])}
    if name=='inspect-existing':return office.inspect_existing(**request)
    if name=='draft':
        from production_app.draft_package import prepare_draft
        return prepare_draft(**request)
    if name=='check':return core.plan(**request)
    if name=='prepare':return core.prepare(**request)
    if name=='verify':return core.verify(**request)
    if name=='preview':return core.preview(**request)
    if name=='confirm':return core.confirm(**request)
    if name=='release':return core.release(**request)
    if name=='commit':return core.commit(**request)
    if name=='readback':return core.readback(**request)
    if name=='complete':return core.complete(**request)
    if name=='recover':return core.recover(**request)
    raise RuntimeError('Неизвестная операция: '+name)

def main():
    from production_app.storage import read,save
    p=argparse.ArgumentParser();p.add_argument('--task');p.add_argument('--request');p.add_argument('--result');p.add_argument('--office-worker')
    args=p.parse_args()
    if not args.task:
        from production_app.gui import launch
        launch();return 0
    if not args.office_worker:
        from stage9_app.supervisor import supervise
        return supervise(sys.argv[1:],entrypoint=Path(__file__))
    from stage9_app.ownership import owned_office
    result=None
    try:
        with owned_office(Path(args.office_worker),supervised=True):
            try:result=task(args.task,read(args.request) if args.request else {})
            except BaseException as exc:
                # Convert the exception to plain text BEFORE exiting the COM
                # guard. Traceback frames otherwise retain native references.
                result={'status':'FAIL','error':str(exc),'traceback':traceback.format_exc()}
    except BaseException as exc:
        if isinstance(result,dict) and result.get('status')=='FAIL':
            # Retain the original operation failure when teardown also fails.
            # Both remain blocking; the supervisor still validates ownership.
            result={**result,'cleanup_error':str(exc),
                    'cleanup_traceback':traceback.format_exc()}
        else:result={'status':'FAIL','error':str(exc),'traceback':traceback.format_exc()}
    save(args.result,result)
    return 0 if result.get('status')=='PASS' else 1

if __name__=='__main__':
    if '--office-cleanup' in sys.argv:
        from stage9_app.supervisor import cleanup_entry
        sys.exit(cleanup_entry(sys.argv[sys.argv.index('--office-cleanup')+1]))
    sys.exit(main())
