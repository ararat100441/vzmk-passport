from common import *
parent=Path(SETUP['parent'])/'unit_fixtures';parent.mkdir(exist_ok=True)
env=os.environ.copy();env['VZMK_TEST_PARENT']=str(parent);env['PYTHONUTF8']='1'
results=[]
for name,relative,count in [('AKZ_AUTO_32','tests/stage9_akz_auto/test_auto.py',32),('OFFICE_REGRESSION_17','tests/stage9_098/test_cleanup.py',17)]:
    log=OUT/(name+'.log')
    with log.open('xb') as f:p=subprocess.run([str(PY),'-B',str(ROOT/relative)],env=env,stdout=f,stderr=subprocess.STDOUT,timeout=180)
    content=log.read_text(encoding='utf8');assert p.returncode==0 and f'Ran {count} tests' in content and '\nOK' in content,content
    results.append({'status':'PASS','runtime':'PYTHON_SOURCE_ONLY','tests':count,'test':str(ROOT/relative),'sha256':sha(ROOT/relative),'log':str(log)})
save(OUT/'SOURCE_TESTS_FINAL.json',{'status':'PASS','results':results,'GUI':'NOT_RUN','native_Office':'RECORDED_SEPARATELY'})
print('Source AKZ 32/32; Office regression 17/17 PASS',flush=True)
