from common import *
p=OUT/'foreign_office.py';s=p.read_text(encoding='utf8')
s=s.replace("RUN=OUT/('foreign_office_'+uuid4().hex[:8])","RUN=OUT/('foreign_pumped_'+uuid4().hex[:8])")
inject='''
# Keep the sentinel owner's STA responsive while a separate frozen job runs.
def run_pumped(command,env=None,timeout=None,capture_output=False,**kw):
    proc=subprocess.Popen(command,env=env,stdout=subprocess.PIPE if capture_output else None,stderr=subprocess.PIPE if capture_output else None,**kw)
    deadline=time.monotonic()+(timeout or 450)
    while proc.poll() is None:
        pythoncom.PumpWaitingMessages()
        if time.monotonic()>deadline:raise TimeoutError('PUMPED_CONTROLLER_TIMEOUT')
        time.sleep(.02)
    out,err=proc.communicate()
    return subprocess.CompletedProcess(command,proc.returncode,out,err)
subprocess.run=run_pumped
'''
s=s.replace('pythoncom.CoInitialize()','pythoncom.CoInitialize()'+inject,1)
(OUT/'foreign_pumped.py').write_text(s,encoding='utf8')
save(OUT/'FOREIGN_CONTROLLER_DIAGNOSTIC_PLAN.json',{'original_attempt_preserved':str(next(OUT.glob('foreign_office_*'))),'changed_only_controller':'pump the sentinel owner STA during child waits; no concurrent GUI teardown','product_changes':False,'controller_sha256':sha(OUT/'foreign_pumped.py'),'purpose':'separate harness blocking from frozen reliability; previous FAIL remains recorded'})
