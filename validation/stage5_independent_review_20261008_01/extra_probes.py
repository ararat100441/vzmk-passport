import json, os, sys, tempfile
from copy import deepcopy
from pathlib import Path
from openpyxl import Workbook
from unittest.mock import patch
RUN=Path(__file__).resolve().parent;ROOT=RUN.parents[1]
sys.path.insert(0,str(ROOT));tempfile.tempdir=str(RUN/'tmp')
import uzk_stage5.inputs as I
import uzk_stage5.core as C
from profiles_stage4.offline import no_network
from profiles_stage4.model import ProfileError
import socket,win32file,win32con
import faulthandler
trace=(RUN/'extra_probe_stack.log').open('w',encoding='utf8')
faulthandler.dump_traceback_later(15,repeat=True,file=trace)
checks=[]
def check(name,fn):
    print('START '+name,flush=True)
    try:checks.append(dict(name=name,status='PASS',detail=fn()))
    except Exception as e:checks.append(dict(name=name,status='FAIL',error=str(e)))
    print('END '+name+' '+checks[-1]['status'],flush=True)
def invalid_registry():
    results=[]
    for i,value in enumerate((0,-1,1.5,True,'=2+2','NaN','Infinity','bad')):
        path=RUN/'fixtures'/f'invalid_registry_{i}.xlsx';b=Workbook();s=b.active;s.title='УЗК'
        s.append(['Заказ','Марка элемента','Кол-во (шт)']);s.append(['PROJECT','MARK',value]);b.save(path);b.close();before=I.sha(path)
        try:I.registry_rows(path)
        except I.Stage5Error as e:assert 'QUANTITY_INVALID' in str(e);results.append(str(e))
        else:raise AssertionError('Invalid registry accepted: '+repr(value))
        assert I.sha(path)==before
    return results
check('invalid_quantity_in_actual_fixture_workbook_reader',invalid_registry)
def actual_locked_existing():
    data=json.loads((RUN/'fresh_preview.json').read_text(encoding='utf8'))['data'];data=deepcopy(data);data['passport']['number']='888/2099'
    p=C.Preview(data,C.digest(data));folder=RUN/'fixtures/locked_existing';folder.mkdir(exist_ok=True);out=folder/'888_2099 - УЗК (ТЕСТ).docx';out.write_bytes(b'LOCKED EXISTING FIXTURE')
    before=I.sha(out);handle=win32file.CreateFile(str(out),win32con.GENERIC_READ,0,None,win32con.OPEN_EXISTING,0,None)
    try:
        try:C.generate(p,folder)
        except I.Stage5Error as e:assert 'OUTPUT_EXISTS' in str(e);error=str(e)
        else:raise AssertionError('Locked existing output was accepted')
    finally:win32file.CloseHandle(handle)
    assert I.sha(out)==before
    return dict(error=error,existing_bytes_unchanged=True,lock='native Windows CreateFile share=0')
check('native_Windows_locked_existing_output_preserved',actual_locked_existing)
def network_denied():
    with no_network() as attempts:
        try:socket.getaddrinfo('fixture.invalid',80)
        except ProfileError:pass
        else:raise AssertionError('Guard did not deny resolver')
    assert len(attempts)==1;return attempts
check('network_denial_before_any_request',network_denied)
def empty_preview_produces_no_output():
    cat=ROOT/'validation/stage5_independent_review_20261008_01/scenarios/no_uzk_True/Документы/Гараж казахстан/Справочник марок.xlsx'
    writes=[];active=[True]
    def audit(event,args):
        if event=='open' and active[0] and isinstance(args[0],(str,bytes)):
            mode=args[1];flags=args[2]
            if (isinstance(mode,str) and any(c in mode for c in 'wax+')) or (isinstance(flags,int) and flags & (os.O_WRONLY|os.O_RDWR|os.O_CREAT|os.O_TRUNC)):
                writes.append(str(args[0]))
    sys.addaudithook(audit)
    try:
        with patch.object(C,'project_sources',return_value=(cat.parent,cat,cat.parent/'Паспорт/1_2026 - УЗК.docx')):
            try:C.preview('Гараж Казахстан','9/2026',cat.parent/'source.xls')
            except I.Stage5Error as e:assert 'MARK_PROFILE_CONFLICT' in str(e)
            else:raise AssertionError('Conflict unexpectedly accepted')
    finally:active[0]=False
    assert writes==[];return dict(full_preview_failed=True,write_events=writes)
check('failed_preview_no_output',empty_preview_produces_no_output)
(RUN/'extra_probes.json').write_text(json.dumps(checks,ensure_ascii=False,indent=2),encoding='utf8')
print(json.dumps(checks,ensure_ascii=False,indent=2))
faulthandler.cancel_dump_traceback_later()
