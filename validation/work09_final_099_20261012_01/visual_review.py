from common import *
run=next(OUT.glob('e2e_*'))
result=min((p for p in run.glob('generate_*.json') if not p.name.endswith(('.command.json','.process.json'))),key=lambda p:p.stat().st_ctime_ns)
m=read(read(result)['manifest']);pages=m['native_quality']['images']
save(OUT/'VISUAL_WORD_REVIEW.json',{'status':'PASS_OBSERVED_PAGES','manifest':read(result)['manifest'],
    'images':{p:sha(p) for p in pages},'pages_inspected':2,'passport':'9/2026',
    'observations':['TEST draft official issue prohibition visible','number 9/2026','mass 15822.14 kg',
    'manufacturing dates 05.08.2026 and 06.10.2026','both VIK 9/2026 and 9/2026-1 dated 29.09.2026',
    'UZK 01.10.2026 and AKZ 05.10.2026 listed','points 1-14 readable without observed clipping','manual signature blank'],
    'all_other_pages_visual_review':'NOT_RUN; native content/normal-open checks separate'})
