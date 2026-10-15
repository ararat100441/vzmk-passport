from pathlib import Path
import os,shutil
from stage8_closure import core,storage as S,native
from stage8_closure.fixtures import create
from support import RUN,reg,partial,stopped,proc,save,states,balances
if __name__=='__main__':
    seed=create(RUN/'sandboxes');S.write(RUN/'SEED.json',seed)
    r,crash=partial();op=S.history(r)['active'];old_id=S.file_identity(reg(r));h=S.sha(reg(r))
    first=stopped(lambda:core.recover(r),'AMBIGUOUS_OPERATION')
    assert not hasattr(core,'restore')
    fresh=proc('recover',r);assert fresh['exit']==2,fresh
    damaged=r/'DAMAGED_EVIDENCE.bin';shutil.copyfile(reg(r),damaged)
    source=r/'transactions/manual_verified_backup.xlsm';shutil.copyfile(op['backup']['path'],source);os.replace(source,reg(r))
    second=stopped(lambda:core.recover(r),'FILE_ID');assert S.file_identity(reg(r))!=old_id
    assert balances(r)=={'s8-a':12,'s8-b':10} and 'COMMITTED' not in states(r)
    S.write(RUN/'BEFORE_REPRODUCTION.json',{'status':'CONFIRMED','sandbox':str(r),'crash':crash,'partial_sha':h,'partial_stop':first,'fresh_process_stop':fresh,'manual_backup_identity_stop':second,'old_id':old_id,'new_id':S.file_identity(reg(r)),'backup_native':native.inspect(op['backup']['path'],{'УЗК!C24':12}),'journal':states(r)})
    print('BEFORE_DEFECT_CONFIRMED',flush=True)
