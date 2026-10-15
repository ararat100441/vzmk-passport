import sys,json
from pathlib import Path
RUN=Path(__file__).resolve().parent;sys.path.insert(0,str(RUN.parents[1]));sys.dont_write_bytecode=True
from stage8_closure import storage as S,model
import psutil
a=S.read(RUN/'GUI_01_PENDING.json');b=S.read(RUN/'GUI_03_CANCEL.json');c=S.read(RUN/'GUI_04_RECOVERED.json');d=S.read(RUN/'GUI_05_FRESH_PREVIEW.json')
assert a==b and a['close']=='disabled' and a['restore']=='normal' and not a['confirmation']
assert S.read(RUN/'GUI_CANCEL_UNCHANGED.json')['PASS']
assert 'RECOVERED_ABORTED' in c['text'] and 'Списаний: 0' in c['text'] and c['plan'] is None and c['restore_plan'] is None and c['close']=='disabled' and c['restore']=='disabled'
assert d['plan']['payload']['status']=='READY_TEST' and d['close']=='normal' and not d['confirmation'] and d['restore']=='disabled'
assert [(x['before'],x['quantity'],x['after']) for x in d['plan']['payload']['deductions']]==[(12,3,9),(10,4,6)]
r=Path(d['sandbox']);h=S.history(r);assert h['active'] is None and not h['closed'] and len(h['restored'])==1
op=h['restored'][0];assert S.sha(r/'registry/УЗК_TEST.xlsm')==op['backup']['sha256']
assert S.file_identity(r/'registry/УЗК_TEST.xlsm')==op['restored_identity']!=op['plan']['registry_identity']
assert not psutil.pid_exists(S.read(RUN/'GUI_PROCESS.json')['pid'])
q={'status':'PASS','visible_real_Tk_App':True,'input':'Computer Use @oai/sky; actual file chooser, buttons, default NO modal, Return cancel and Yes restore; no messagebox/handler patches','sandbox':str(r),'restore_verified_stage_length':len(op['restoration']['stage']),'default_NO_and_cancel_unchanged':True,'RECOVERED_ABORTED_zero_decrements':True,'lawful_new_identity':op['restored_identity'],'no_COMMITTED':True,'fresh_preview_12_3_9_and_10_4_6':True,'confirmation_reset':True,'own_GUI_process_exited':True,'manual_human_User_E2E':'NOT_RUN'}
(RUN/'GUI_VALIDATION.json').write_text(json.dumps(q,ensure_ascii=False,indent=2),encoding='utf8');print('GUI_PASS')
