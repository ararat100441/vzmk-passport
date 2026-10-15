from common import *
from decimal import Decimal,localcontext
from stage9_app import akz_auto as A
from stage8_closure.storage import Stop
import math,subprocess
env=os.environ.copy();env['VZMK_TEST_PARENT']=str(Path(config()['base'])/'source_tests');env['PYTHONUTF8']='1'
p=subprocess.run([PY,'-B',str(ROOT/'tests/stage9_akz_auto/test_auto.py')],env=env,capture_output=True,timeout=90)
with (OUT/'SOURCE_32.log').open('xb') as f:f.write(p.stdout+p.stderr)
assert p.returncode==0
vectors=[]
with localcontext() as ctx:
 ctx.prec=50
 for t in (-45,-10,0,15,17.5,20,23.1,25,60):
  for h in (1,40,40.1,50,60,64.9,65,100):
   td=Decimal(str(t));rh=Decimal(str(h));a=Decimal('17.62');b=Decimal('243.12');g=(rh/100).ln()+a*td/(b+td);ref=b*g/(a-g)
   valid=Decimal(-45)<=ref<=Decimal(60)
   try:actual=A.dew_point(t,h)
   except Stop:
    assert not valid,(t,h,ref);vectors.append({'t_C':t,'RH_percent':h,'expected':'DOMAIN_STOP','reference':str(ref)});continue
   assert valid and abs(actual-float(ref))<2e-12,(t,h,ref,actual)
   vectors.append({'t_C':t,'RH_percent':h,'reference':str(ref),'actual':actual,'absolute_error':abs(actual-float(ref))})
for t,h in [(293.15,50),(20,0),(20,101),(20,float('nan')),(float('inf'),50),(True,50),(20,'50')]:
 try:A.dew_point(t,h)
 except Stop:pass
 else:raise AssertionError(('INVALID_ACCEPTED',t,h))
folder=Path(config()['base'])/'math_days';folder.mkdir();cfg={'air_c':[15,25],'surface_c':[18,28],'relative_humidity':[40,65],'formula':A.FORMULA}
samples=[]
from datetime import date,timedelta
for n in range(128):
 r=A.daily_climate(folder,(date(2031,1,1)+timedelta(days=n)).isoformat(),cfg);v=r['values'];ref=A.dew_point(v['air_c'],v['relative_humidity'])
 assert 15<=v['air_c']<=25 and 18<=v['surface_c']<=28 and 40<=v['relative_humidity']<=65
 assert v['dew_point_c']==round(ref,1) and v['surface_c']>=max(ref,v['dew_point_c'])+3-1e-10
 assert all(abs(v[k]*10-round(v[k]*10))<1e-10 for k in ('air_c','surface_c','relative_humidity','dew_point_c'))
 samples.append(v)
save(OUT/'SOURCE_AND_MAGNUS.json',{'status':'PASS','source_only':True,'existing_targeted_tests':32,'independent_Decimal50_vectors':vectors,'invalid_input_cases':7,'random_daily_samples':samples,'formula':'ln(RH_percent/100)+17.62*T_C/(243.12+T_C)','WMO_source':'https://etrp.wmo.int/pluginfile.php/97653/mod_folder/content/0/Training%20Material/5_1_Humidity%20Part%201%20Introduction.pdf#page=13','Celsius_and_percent':True,'water_domain_C':[-45,60],'RH_domain_percent':'0<H<=100','display_precision_C':0.1,'dew_margin_checked_against_raw_and_rounded':True,'GUI_created':False})
print('SOURCE32 + Decimal50 Magnus72 vectors + invalid7 +128 climates PASS',flush=True)
