"""Actual EXE fault test: reserve only its own final diagnostic, never overwrite."""
from common import *
RUN=OUT/('diagnostic_collision_'+uuid4().hex[:8]);RUN.mkdir();cases=[]
diagnostics=Path(SETUP['base'])/'VZMKPassport/diagnostics'
workspace=read(OUT/'LIFECYCLE_FINAL.json')['workspace']
for mode in ('normal','cleanup-error'):
    result=RUN/('RESULT_'+mode+'.json')
    command=[str(EXE),'--task','office-probe','--probe-mode',mode,'--workspace',workspace,'--result',str(result)]
    p=subprocess.Popen(command,env=ENV,creationflags=subprocess.CREATE_NO_WINDOW,stdout=subprocess.PIPE,stderr=subprocess.PIPE)
    deadline=time.monotonic()+20;job=None
    while time.monotonic()<deadline and p.poll() is None:
        for request in diagnostics.glob('job_*/JOB_REQUEST.json'):
            try:d=read(request)
            except (OSError,ValueError):continue
            if d['supervisor_pid']==p.pid and d['supervisor_image'].casefold()==str(EXE).casefold() and str(result) in d['arguments']:
                job=request.parent;break
        if job is not None:break
        time.sleep(.01)
    assert job is not None,'OWN_JOB_NOT_IDENTIFIED'
    final=job/'office.json';save(final,{'test_only_reserved_evidence':True,'mode':mode,'supervisor_pid':p.pid})
    before=sha(final);out,err=p.communicate(timeout=450)
    save(RUN/('PROCESS_'+mode+'.json'),{'command':command,'returncode':p.returncode,'stdout':out.decode('utf8','replace'),'stderr':err.decode('utf8','replace')})
    data=read(result);sup=read(data['office_supervisor'])
    assert p.returncode==1 and data['status']=='FAIL' and sha(final)==before
    assert sup['owned_orphans']==0 and not sup['ownership_errors'] and len(sup['instances'])==2
    assert all(r['exit_code']==0 and not r.get('forced_termination_attempted') for r in sup['instances'])
    serialized=json.dumps(data,ensure_ascii=False);assert 'Diagnostic.FinalSave' in serialized and 'File exists' in serialized
    if mode=='cleanup-error':
        assert data['error']=='STAGE9_TEST_OFFICE_ERROR'
        assert 'STAGE9_TEST_WORD_CLOSE_ERROR' in serialized and 'Word.Close' in serialized
    assert (job/'office_FAILURE.json').is_file()
    cases.append({'mode':mode,'job':str(job),'result':str(result),'supervisor':data['office_supervisor'],
        'reserved_final':str(final),'reserved_final_sha256':before,'old_evidence_preserved':True,
        'expected_status':'FAIL','native_office_exit_codes':[r['exit_code'] for r in sup['instances']],
        'primary_preserved':mode=='cleanup-error','worker_failure_evidence':str(job/'office_FAILURE.json')})
    print('Frozen diagnostic collision',mode,'safe STOP PASS; evidence preserved; Office exit 0',flush=True)
save(OUT/'OFFICE_DIAGNOSTIC_COLLISION_FINAL.json',{'status':'PASS','exe_sha256':sha(EXE),'cases':cases,'GUI':False})
