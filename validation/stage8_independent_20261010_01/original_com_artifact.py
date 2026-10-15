import independent as T
from pathlib import Path
from stage8_closure import storage as S,native
case=S.read(T.ROOT/'validation/stage8_20261010_03/targeted/25_external_registry_writer_is_denied.json')
root=Path(case['detail']['sandbox']);registry=T.reg(root);meta=S.read(root/'SANDBOX.json');before=S.sha(registry)
info=native.inspect(registry,{'УЗК!C24':12,'УЗК!C25':10})
assert before==S.sha(registry)==meta['initial_registry_sha256']
assert info['vba_streams']==meta['native_seed']['vba_streams'] and len(info['vba_streams'])==11
assert info['sheets']==meta['native_seed']['sheets']
assert {k:v['formula'] for k,v in info['formulas'].items()}=={k:v['formula'] for k,v in meta['native_seed']['formulas'].items()}
T.save(T.RUN/'ORIGINAL_COM_AFFECTED_ARTIFACT.json',{'status':'PASS','reported_case':case['name'],'reported_stack':'fixtures.py:48 create -> book.Save',
 'registry':str(registry),'native_read_only':info,'sha_unchanged':True,'sha_matches_seed':True,'vba_streams_identical':11,
 'spontaneous_original_cause':'NOT_ESTABLISHED','own_deliberate_RPC_E_DISCONNECTED':'safe refusal; see cases/real_excel_disconnect_at_save.json'})
print('Original affected fixture: normal read-only native open PASS; 11 VBA streams; balances12/10; SHA matches seed and unchanged.')
