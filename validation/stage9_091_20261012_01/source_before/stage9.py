"""Single executable entry point; workers use the same bundled interpreter."""
import argparse
import json
import os
from pathlib import Path
import sys
sys.dont_write_bytecode=True

def main():
    p=argparse.ArgumentParser()
    p.add_argument('--task',choices=['import','check','generate','fixture','close-preview','close','recover','restore-preview','restore','verify','smoke'])
    p.add_argument('--workspace');p.add_argument('--source');p.add_argument('--parent');p.add_argument('--vom')
    p.add_argument('--project',default='Гараж Казахстан');p.add_argument('--passport',default='9/2026')
    p.add_argument('--variants',default='0');p.add_argument('--new',action='store_true');p.add_argument('--preview')
    p.add_argument('--request');p.add_argument('--sandbox');p.add_argument('--manifest');p.add_argument('--result')
    p.add_argument('--confirm');p.add_argument('--confirm-test',action='store_true')
    args=p.parse_args()
    if args.task is None:
        from stage9_app.gui import launch
        return launch()
    from stage9_app.workspace import bootstrap,bind,save
    from stage8_closure import storage as S
    result={'status':'FAIL'}
    try:
        if args.task=='import':result={'workspace':str(bootstrap(args.source,args.parent))}
        else:
            root,_=bind(args.workspace)
            from stage9_app import core
            from stage8_closure import core as closure,recovery,fixtures
            if args.task=='check':
                plan=core.check(root,args.project,args.passport,args.vom,tuple(int(n) for n in args.variants.split(',') if n!=''),args.new)
                result={'preview':str(Path(plan['data']['run'])/'PREVIEW.json'),'summary':core.describe(plan),'data':plan['data']}
            elif args.task=='generate':
                old=S.read(args.preview);seal=old.pop('seal')
                if Path(old['root']).resolve()!=root or not S.no_links(args.preview).is_relative_to(root/'validation'):
                    raise S.Stop('Предварительный просмотр принадлежит другому DEV-стенду.')
                core.verify({'data':old,'seal':seal})
                plan=core.check(root,old['project'],old['number'],old['source_vom'],old['variants'],old['new'])
                result={'manifest':str(core.generate(plan,args.confirm_test))}
            elif args.task=='fixture':
                result=fixtures.create(root/'validation'/('stage8_s9_'+__import__('uuid').uuid4().hex[:8])/'sandboxes')
            elif args.task=='close-preview':result={'plan':closure.preview(args.sandbox,args.request)}
            elif args.task=='close':result=closure.close(S.read(args.preview),args.confirm or '')
            elif args.task=='recover':
                try:result=closure.recover(args.sandbox)
                except S.Stop as exc:
                    result={'plan':recovery.preview(args.sandbox),'recovery_reason':str(exc)}
            elif args.task=='restore-preview':result={'plan':recovery.preview(args.sandbox)}
            elif args.task=='restore':result=recovery.restore(S.read(args.preview),args.confirm or '')
            elif args.task=='verify':
                m=core.verify_manifest(args.manifest);result={'passport':m['passport'],'documents':len(m['records'])}
            elif args.task=='smoke':
                import tkinter as tk
                from stage9_app.gui import App
                w=tk.Tk();app=App(w,root);w.update();result={'title':w.title(),'version':'0.9','frozen':bool(getattr(sys,'frozen',False)),'python_in_path':__import__('shutil').which('python'),'gui':True};app.shutdown()
        if 'status' in result:result['operation_status']=result.pop('status')
        result={'status':'PASS',**result}
    except Exception as exc:
        import traceback
        result={'status':'FAIL','error':str(exc),'traceback':traceback.format_exc()}
    if args.result:save(args.result,result)
    elif sys.stdout:print(json.dumps(result,ensure_ascii=False,indent=2,default=str))
    return 0 if result['status']=='PASS' else 1

if __name__=='__main__':
    from profiles_stage4.offline import no_network
    # GUI viewers live until explicit app close; CLI Office instances are worker-owned.
    if '--task' in sys.argv:
        from stage9_app.ownership import owned_office
        from stage9_app.workspace import save
        from uuid import uuid4
        diagnostic=Path(os.environ.get('LOCALAPPDATA',str(Path.home())))/'VZMKPassport/diagnostics'/('office_'+uuid4().hex+'.json')
        final=None;temporary=None
        if '--result' in sys.argv:
            slot=sys.argv.index('--result')+1;final=Path(sys.argv[slot])
            temporary=final.with_name('.pending_'+uuid4().hex+'.json');sys.argv[slot]=str(temporary)
        try:
            with no_network(),owned_office(diagnostic):code=main()
            if temporary is not None:os.rename(temporary,final)
        except Exception as exc:
            import traceback
            error={'status':'FAIL','error':str(exc),'traceback':traceback.format_exc(),'phase':'worker_finalization','office_diagnostic':str(diagnostic)}
            save(diagnostic.with_name(diagnostic.stem+'_FAILURE.json'),error)
            if final is not None and not final.exists():save(final,error)
            code=1
        sys.exit(code)
    with no_network():sys.exit(main())
