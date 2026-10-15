"""Single executable entry point; workers use the same bundled interpreter."""
import argparse
import json
import os
from pathlib import Path
import sys
sys.dont_write_bytecode=True
if os.name=='nt':
    # Accepted native rendering calls the Windows PowerShell executable by name.
    # Resolve the OS tool even when the caller supplies a minimal PATH.
    powershell=Path(os.environ.get('SystemRoot',r'C:\Windows'))/'System32/WindowsPowerShell/v1.0'
    if (powershell/'powershell.exe').is_file():os.environ['PATH']=str(powershell)+os.pathsep+os.environ.get('PATH','')

def main():
    p=argparse.ArgumentParser()
    p.add_argument('--task',choices=['import','check','generate','fixture','close-preview','close','recover','restore-preview','restore','verify','native-verify','smoke','office-probe','akz-settings-init','akz-settings-save','akz-auto-check','akz-auto-generate'])
    p.add_argument('--probe-mode',choices=['normal','error','crash','cleanup-error'],default='normal',help=argparse.SUPPRESS)
    p.add_argument('--test-fault',choices=['before_save','save_io','after_save','during_publish'],help=argparse.SUPPRESS)
    p.add_argument('--workspace');p.add_argument('--source');p.add_argument('--parent');p.add_argument('--vom')
    p.add_argument('--project',default='Гараж Казахстан');p.add_argument('--passport',default='9/2026')
    p.add_argument('--variants',default='0');p.add_argument('--new',action='store_true');p.add_argument('--preview')
    p.add_argument('--request');p.add_argument('--sandbox');p.add_argument('--manifest');p.add_argument('--result')
    p.add_argument('--confirm');p.add_argument('--confirm-test',action='store_true')
    p.add_argument('--akz-settings')
    args=p.parse_args()
    if args.task is None:
        from stage9_app.gui import launch
        return launch()
    from stage9_app.workspace import bootstrap,bind,save
    from stage8_closure import storage as S
    result={'status':'FAIL'};cleanup_errors=[]
    try:
        if args.task=='import':result={'workspace':str(bootstrap(args.source,args.parent))}
        else:
            root,_=bind(args.workspace)
            from stage9_app import core
            from stage8_closure import core as closure,recovery,fixtures
            if args.task in ('akz-settings-init','akz-settings-save'):
                from stage9_app import akz_auto
                cfg=akz_auto.defaults(root) if args.task=='akz-settings-init' else S.read(S.no_links(args.akz_settings))
                result=akz_auto.store_settings(root,cfg)
            elif args.task in ('check','akz-auto-check'):
                if args.task=='akz-auto-check':
                    from stage9_app import akz_auto
                    if not akz_auto.enabled(root):raise S.Stop('AKZ_AUTO_DISABLED')
                plan=core.check(root,args.project,args.passport,args.vom,tuple(int(n) for n in args.variants.split(',') if n!=''),args.new)
                result={'preview':str(Path(plan['data']['run'])/'PREVIEW.json'),'summary':core.describe(plan),'data':plan['data']}
            elif args.task in ('generate','akz-auto-generate'):
                old=S.read(args.preview);seal=old.pop('seal')
                if Path(old['root']).resolve()!=root or not S.no_links(args.preview).is_relative_to(root/'validation'):
                    raise S.Stop('Предварительный просмотр принадлежит другому DEV-стенду.')
                core.verify({'data':old,'seal':seal})
                plan=core.check(root,old['project'],old['number'],old['source_vom'],old['variants'],old['new'])
                if args.task=='akz-auto-generate':
                    from stage9_app import akz_auto
                    if not args.confirm_test:raise S.Stop('Подтвердите TEST-моделирование АКЗ.')
                    result={'pairs':akz_auto.generate_groups(root,plan['batch'],plan['data']['variants'],Path(plan['data']['run']).name)}
                else:result={'manifest':str(core.generate(plan,args.confirm_test))}
            elif args.task=='fixture':
                result=fixtures.create(root/'validation'/('stage8_s9_'+__import__('uuid').uuid4().hex[:8])/'sandboxes')
            elif args.task=='close-preview':result={'plan':closure.preview(args.sandbox,args.request)}
            elif args.task=='close':
                plan=S.read(args.preview)
                if args.test_fault:
                    payload=plan['payload']
                    request=S.read(payload['request_path'])
                    if request.get('project')!='STAGE8 ИСКУССТВЕННЫЙ ПРОЕКТ' or request.get('passport')!='8001/2099' or request.get('test_only') is not True:raise S.Stop('FAULT_INJECTION_REQUIRES_ARTIFICIAL_FIXTURE')
                    def fault(point,path):
                        if point==args.test_fault:
                            if point=='during_publish':os._exit(73)
                            raise OSError('STAGE9_TEST_INJECTED_'+point)
                    result=closure.close(plan,args.confirm or '',_fault=fault)
                else:result=closure.close(plan,args.confirm or '')
            elif args.task=='recover':
                try:result=closure.recover(args.sandbox)
                except S.Stop as exc:
                    result={'plan':recovery.preview(args.sandbox),'recovery_reason':str(exc)}
            elif args.task=='restore-preview':result={'plan':recovery.preview(args.sandbox)}
            elif args.task=='restore':result=recovery.restore(S.read(args.preview),args.confirm or '')
            elif args.task=='verify':
                m=core.verify_manifest(args.manifest);result={'passport':m['passport'],**core.document_counts(m)}
            elif args.task=='native-verify':
                from stage9_app.validation import inspect_bundle
                result=inspect_bundle(root,args.manifest)
            elif args.task=='smoke':
                import tkinter as tk
                from stage9_app.gui import App
                from stage9_app import VERSION
                w=tk.Tk();app=App(w,root);w.update();result={'title':w.title(),'version':VERSION,'frozen':bool(getattr(sys,'frozen',False)),'python_in_path':__import__('shutil').which('python'),'gui':True};app.shutdown()
            elif args.task=='office-probe':
                import win32com.client
                excel=word=book=doc=None
                try:
                    excel=win32com.client.DispatchEx('Excel.Application');excel.AutomationSecurity=3;excel.DisplayAlerts=False;excel.EnableEvents=False
                    book=excel.Workbooks.Add();book.Worksheets(1).Cells(1,1).Value2='ARTIFICIAL STAGE9 OFFICE LIFECYCLE TEST'
                    word=win32com.client.DispatchEx('Word.Application');word.AutomationSecurity=3;word.DisplayAlerts=0
                    doc=word.Documents.Add(Visible=False);doc.Content.Text='ARTIFICIAL STAGE9 OFFICE LIFECYCLE TEST'
                    if args.probe_mode=='crash':os._exit(73)
                    if args.probe_mode in ('error','cleanup-error'):raise RuntimeError('STAGE9_TEST_OFFICE_ERROR')
                    result={'probe_mode':args.probe_mode,'excel':True,'word':True}
                finally:
                    primary=sys.exc_info()[1]
                    for name,application,document in (('Word',word,doc),('Excel',excel,book)):
                        for step,action in (('Close',lambda:document.Close(False) if document is not None else None),
                                            ('Quit',lambda:application.Quit() if application is not None else None)):
                            try:
                                action()
                                if name=='Word' and step=='Close' and args.probe_mode=='cleanup-error':raise RuntimeError('STAGE9_TEST_WORD_CLOSE_ERROR')
                            except Exception as exc:cleanup_errors.append({'step':name+'.'+step,'error':repr(exc)})
                    application=document=word=excel=doc=book=None
                    if primary is None and cleanup_errors:raise RuntimeError('OFFICE_PROBE_CLEANUP_FAILED: '+str(cleanup_errors))
        if 'status' in result:result['operation_status']=result.pop('status')
        result={'status':'PASS',**result}
    except Exception as exc:
        import traceback
        result={'status':'FAIL','error':str(exc),'traceback':traceback.format_exc()}
        if cleanup_errors:result['cleanup_errors']=cleanup_errors
    if args.result:save(args.result,result)
    elif sys.stdout:print(json.dumps(result,ensure_ascii=False,indent=2,default=str))
    return 0 if result['status']=='PASS' else 1

if __name__=='__main__':
    from profiles_stage4.offline import no_network
    if '--office-cleanup' in sys.argv:
        from stage9_app.supervisor import cleanup_entry
        with no_network():sys.exit(cleanup_entry(sys.argv[sys.argv.index('--office-cleanup')+1]))
    # GUI viewers live until explicit app close; CLI Office instances are worker-owned.
    if '--task' in sys.argv:
        if '--office-worker' not in sys.argv:
            from stage9_app.supervisor import supervise
            with no_network():sys.exit(supervise(sys.argv[1:]))
        from stage9_app.ownership import owned_office
        from stage9_app.workspace import save
        from uuid import uuid4
        worker_slot=sys.argv.index('--office-worker');diagnostic=Path(sys.argv[worker_slot+1]);del sys.argv[worker_slot:worker_slot+2]
        final=None;temporary=None
        if '--result' in sys.argv:
            slot=sys.argv.index('--result')+1;final=Path(sys.argv[slot])
            temporary=final.with_name('.pending_'+uuid4().hex+'.json');sys.argv[slot]=str(temporary)
        try:
            with no_network(),owned_office(diagnostic,supervised=True):code=main()
            if temporary is not None:os.rename(temporary,final)
        except Exception as exc:
            import traceback
            error={'status':'FAIL','error':str(exc),'traceback':traceback.format_exc(),'phase':'worker_finalization','office_diagnostic':str(diagnostic)}
            if temporary is not None and temporary.exists():
                from stage8_closure.storage import read
                prior=read(temporary);error['worker_result']=prior
                if prior.get('status')=='FAIL':
                    error['finalization_error']=error['error'];error['error']=prior['error']
                    if prior.get('cleanup_errors'):error['cleanup_errors']=prior['cleanup_errors']
            save(diagnostic.with_name(diagnostic.stem+'_FAILURE.json'),error)
            if final is not None and not final.exists():save(final,error)
            code=1
        sys.exit(code)
    with no_network():sys.exit(main())
