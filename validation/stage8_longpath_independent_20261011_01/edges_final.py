"""Fresh independent boundary tests. No product edits; TEST copies only."""
import sys,os,json,shutil,hashlib,time,traceback,gc,subprocess
from pathlib import Path
from uuid import uuid4
from unittest.mock import patch
from contextlib import contextmanager
RUN=Path(__file__).resolve().parent;ROOT=RUN.parents[1]
sys.path.insert(0,str(ROOT));sys.dont_write_bytecode=True
from stage8_closure import shortcopy as SC,storage as S,native,office,core,recovery
import review as R,baseline_helpers as H
import win32security,win32api,win32con,pywintypes
OUT=RUN/'edges_final';OUT.mkdir(exist_ok=True)
H.SEED=Path(S.read(RUN/'SEED.json')['sandbox']);H.EXTERNAL=RUN/'EXTERNAL_VALID_4.xlsm'
PRODUCT=S.read(RUN/'PRODUCT_BEFORE.json');RESULTS=[]
def save(p,q):p.write_text(json.dumps(q,ensure_ascii=False,indent=2),encoding='utf8')
def case(name,fn):
    print('START',name,flush=True);t=time.monotonic()
    try:q={'name':name,'status':'PASS','detail':fn()}
    except Exception as e:q={'name':name,'status':'FAIL','error':str(e),'traceback':traceback.format_exc()}
    gc.collect();q['seconds']=round(time.monotonic()-t,3);RESULTS.append(q)
    save(OUT/(name+'.json'),q);save(OUT/'RESULTS.json',{'PASS':sum(x['status']=='PASS' for x in RESULTS),'FAIL':sum(x['status']=='FAIL' for x in RESULTS),'results':RESULTS})
    assert all(S.sha(p)==h for p,h in PRODUCT.items())
    print(q['status'],name,q.get('error',''),flush=True)
def source():
    p=OUT/(uuid4().hex+'.xlsm');shutil.copyfile(H.reg(H.SEED),p);return p
def fail(fn,contains=None):
    try:fn()
    except Exception as e:
        if contains:assert contains in str(e),(contains,str(e))
        return office._error(e)
    raise AssertionError('EXPECTED_REFUSAL')
def permissions_and_lock():
    p=source();before=S.sha(p);attempts=[]
    with SC.workspace() as w:
        folder=w.path;d=win32security.GetFileSecurity(str(folder),win32security.DACL_SECURITY_INFORMATION)
        assert d.GetSecurityDescriptorControl()[0]&win32security.SE_DACL_PROTECTED
        user=win32security.GetTokenInformation(win32security.OpenProcessToken(win32api.GetCurrentProcess(),win32con.TOKEN_QUERY),win32security.TokenUser)[0]
        allowed={'S-1-5-18',win32security.ConvertSidToStringSid(user)}
        acl=d.GetSecurityDescriptorDacl();aces=[acl.GetAce(i) for i in range(acl.GetAceCount())]
        assert len(aces)==2 and {win32security.ConvertSidToStringSid(a[2]) for a in aces}==allowed
        with SC.exact_copy(p,w) as (c,b):
            assert c.read_bytes()==p.read_bytes() and b['source_sha256']==before
            for dst in (p,c):
                for mode in ('write','replace','native'):
                    inc=OUT/(uuid4().hex+'.xlsm');shutil.copyfile(H.EXTERNAL,inc)
                    a=H.proc('publisher',mode,inc,dst);assert a['exit']==2,a
                    attempts.append({'target':str(dst),'mode':mode,'result':a})
            inc=OUT/(uuid4().hex+'.xlsm');shutil.copyfile(H.EXTERNAL,inc)
            a=H.proc('publisher','parent',inc,folder);assert a['exit']==2,a;attempts.append(a)
            assert S.sha(c)==before and S.sha(p)==before
        assert b['source_and_copy_rechecked'] is True and not c.exists()
    assert not folder.exists()
    return {'binding':b,'protected_DACL':True,'only_current_user_and_SYSTEM':True,'attempts':attempts,'cleanup':True}
def long_native(length):
    p=OUT/('long_'+str(length));p.mkdir()
    while len(str(p/'book.xlsm'))<length:
        left=length-len(str(p/'book.xlsm'));p=p/('x'*min(left-1,70) if left>1 else 'x');p.mkdir()
    p=p/'book.xlsm';shutil.copyfile(H.reg(H.SEED),p)
    h=S.sha(p);q=native.inspect(p,{'УЗК!C24':12,'УЗК!C25':10},diagnostics=OUT/'office')
    b=q['verified_copy'];assert b['source_and_copy_rechecked'] and len(b['verification_path'])<=180 and S.sha(p)==h
    assert not Path(b['verification_path']).parent.exists()
    return {'length':len(str(p)),'native':q,'source_unchanged':True,'cleanup':True}
def content_before():
    p=source();folders=[];real=SC.Workspace.track;called=[]
    def track(w,c):
        real(w,c);folders.append(w.path)
        if c.name=='v.xlsm':c.write_bytes(b'ALTERED_TEST_COPY')
    with patch.object(SC.Workspace,'track',track),patch.object(native,'_inspect',side_effect=lambda *a,**k:called.append(a)):
        e=fail(lambda:native.inspect(p),'NATIVE_SHORT_COPY_CHANGED')
    assert not called and all(not p.exists() for p in folders)
    return {'STOP':e,'native_not_called':True,'cleanup':True}
def identity_before():
    p=source();real=SC.Workspace.track;foreign=[];called=[]
    def track(w,c):
        real(w,c)
        if c.name=='v.xlsm':
            t=w/'same_bytes.xlsm';t.write_bytes(c.read_bytes());os.replace(t,c);foreign.append(c)
    with patch.object(SC.Workspace,'track',track):e=fail(lambda:native.inspect(p,diagnostics=OUT/'office'),'NATIVE_CLEANUP_IDENTITY_CHANGED')
    assert foreign and foreign[0].exists() and foreign[0].read_bytes()==p.read_bytes()
    # Keep foreign TEST bytes as evidence; cleanup must never delete them.
    return {'STOP':e,'foreign_preserved':str(foreign[0]),'sha256':S.sha(foreign[0])}
def after_check(kind):
    p=source();realsha=S.sha;realid=S.file_identity;seen=[]
    realinspect=native._inspect
    def inspect(c,*a,**k):
        q=realinspect(c,*a,**k);seen.append(str(c));return q
    def sha(q):
        if seen and ((kind=='copy_sha' and str(q)==seen[0]) or (kind=='source_sha' and Path(q)==p)):return '0'*64
        return realsha(q)
    def identity(q):
        if seen and ((kind=='copy_identity' and str(q)==seen[0]) or (kind=='source_identity' and Path(q)==p)):
            i=realid(q);return i|{'file_index':-1}
        return realid(q)
    with patch.object(native,'_inspect',inspect),patch.object(S,'sha',sha),patch.object(S,'file_identity',identity):
        e=fail(lambda:native.inspect(p,diagnostics=OUT/'office'),'CHANGED')
    assert seen and not Path(seen[0]).parent.exists()
    return {'STOP':e,'real_Excel_completed_before_injected_observation':True,'cleanup':True,'injection':kind}
def ordinary_exception():
    p=source();seen=[]
    def bad(c,*a,**k):seen.append(c);raise RuntimeError('INDEPENDENT_NATIVE_PROBE_FAILURE')
    with patch.object(native,'_inspect',bad):e=fail(lambda:native.inspect(p),'INDEPENDENT_NATIVE_PROBE_FAILURE')
    assert seen and not seen[0].parent.exists()
    return {'exception':e,'cleanup':True}
def directory_collision():
    r=H.fixture();p=R.pclose(r);before=R.snap(r)
    with SC.workspace() as w:
        from types import SimpleNamespace
        with patch.object(SC,'uuid4',return_value=SimpleNamespace(hex=w.path.name[4:])):
            e=fail(lambda:core.close(p,core.confirmation(p)),'REGISTRY_GUARD_OR_IO_FAILURE')
        assert w.path.exists() and not list(w.path.iterdir()) and R.snap(r)==before and not H.states(r)
    assert not w.path.exists()
    return {'STOP':e,'existing_workspace_not_reused':True,'no_INTENT':True,'owned_workspace_cleanup':True}
def collision_foreign():
    p=source();foreign=[]
    def execute():
        with SC.workspace() as w:
            c=w/'v.xlsm';c.write_bytes(b'FOREIGN_EXCLUSIVE_TEST');foreign.append(c)
            with SC.exact_copy(p,w):pass
    e=fail(execute)
    assert foreign and foreign[0].read_bytes()==b'FOREIGN_EXCLUSIVE_TEST'
    return {'STOP':e,'foreign_preserved':str(foreign[0]),'exclusive_create':True}
def guard_wrong():
    r=H.fixture();other=source();before=R.snap(r)
    with S.registry_guard(H.reg(r)) as g:
        with SC.workspace() as w:e=fail(lambda:native.inspect(other,guard=g,verification=w),'GUARD_PATH_MISMATCH')
    assert R.snap(r)==before
    return {'STOP':e,'unchanged':True}
def no_preflight_mutation(kind):
    r=H.fixture();plan=R.pclose(r);before=R.snap(r);called=[]
    if kind=='native_failure':
        ctx=patch.object(native,'_inspect',side_effect=RuntimeError('INDEPENDENT_NATIVE_PREFLIGHT_FAILURE'))
    else:
        base=OUT/('missing_'+uuid4().hex)
        if kind=='file':base.write_bytes(b'NOT_DIRECTORY')
        elif kind=='relative':base=Path('relative_TEST_shortbase')
        elif kind=='long':base=OUT/('z'*90);base.mkdir()
        ctx=patch.object(SC,'BASE',base)
    with ctx:e=fail(lambda:core.close(plan,core.confirmation(plan)))
    assert R.snap(r)==before and not H.states(r) and not list((r/'backups').rglob('*.xlsm'))
    return {'sandbox':str(r),'STOP':e,'no_INTENT_or_backup_or_registry_write':True}
def restore_preflight(resume):
    r,_=R.partial()
    if resume:
        p=R.prest(r);f=r/'RESTORE.json';R.save(f,p);assert R.proc('restore_crash',f,'after_restore_quarantine')['exit']==73
    p=R.prest(r);before=R.snap(r);files={str(f):S.sha(f) for f in (r/'transactions').rglob('*.xlsm')};ids={str(f):S.file_identity(f) for f in (r/'transactions').rglob('*.xlsm')}
    with patch.object(SC,'BASE',Path('relative_TEST_unavailable')):e=fail(lambda:recovery.restore(p,recovery.confirmation(p)),'SHORT_WORKSPACE_UNAVAILABLE')
    assert R.snap(r)==before and all(S.sha(f)==h for f,h in files.items()) and all(S.file_identity(f)==i for f,i in ids.items())
    return {'sandbox':str(r),'resume':resume,'STOP':e,'namespace_and_journal_unchanged':True}
def acl_denied():
    token=win32security.OpenProcessToken(win32api.GetCurrentProcess(),win32con.TOKEN_QUERY)
    try:sid=win32security.ConvertSidToStringSid(win32security.GetTokenInformation(token,win32security.TokenUser)[0])
    finally:token.Close()
    base=Path(os.environ['USERPROFILE'])/'AppData/Local'/('s8ir_'+uuid4().hex)
    sd=win32security.ConvertStringSecurityDescriptorToSecurityDescriptor('D:P(D;;0x00000006;;;'+sid+')(A;OICI;FA;;;SY)(A;OICI;FA;;;'+sid+')',1)
    sa=pywintypes.SECURITY_ATTRIBUTES();sa.SECURITY_DESCRIPTOR=sd
    import win32file
    win32file.CreateDirectory(str(base),sa)
    try:
        r=H.fixture();p=R.pclose(r);before=R.snap(r)
        with patch.object(SC,'BASE',base):e=fail(lambda:core.close(p,core.confirmation(p)),'REGISTRY_GUARD_OR_IO_FAILURE')
        assert R.snap(r)==before and not H.states(r) and not list(base.iterdir())
    finally:
        allow=win32security.ConvertStringSecurityDescriptorToSecurityDescriptor('D:P(A;OICI;FA;;;SY)(A;OICI;FA;;;'+sid+')',1)
        win32security.SetFileSecurity(str(base),win32security.DACL_SECURITY_INFORMATION,allow);base.rmdir()
    return {'sandbox':str(r),'real_denied_DACL':True,'STOP':e,'no_INTENT':True,'base_removed':True}
def independent_native_styles():
    r=H.fixture()
    def observe(p):
        with SC.workspace() as w,SC.exact_copy(p,w) as (c,binding):
            with office.excel_session(diagnostics=OUT/'style_office',workbooks=[c]) as excel:
                book=excel.Workbooks.Open(Filename=str(c.resolve()),UpdateLinks=0,ReadOnly=True,CorruptLoad=0,AddToMru=False)
                try:
                    assert book.ReadOnly and book.HasVBProject and book.FileFormat==52 and excel.AutomationSecurity==3
                    formula=book.Worksheets('S8_FORMULAS');cell=formula.Range('A1')
                    q={'styles_count':int(book.Styles.Count),'cell_style':str(cell.Style),'font_bold':bool(cell.Font.Bold),'font_size':cell.Font.Size,'number_format':cell.NumberFormat,'column_width':cell.ColumnWidth,'formula':cell.Formula,'value':cell.Value2,'sheets':[{'name':s.Name,'visible':int(s.Visible)} for s in book.Worksheets],'links':book.LinkSources(1),'connections':int(book.Connections.Count),'names':[(n.Name,n.RefersTo) for n in book.Names],'C24':book.Worksheets('УЗК').Range('C24').Value2,'C25':book.Worksheets('УЗК').Range('C25').Value2}
                finally:book.Close(SaveChanges=False);book=None
            return q,binding
    before,bb=observe(H.reg(r));c=R.doclose(r);after,ab=observe(H.reg(r))
    assert before['value']==30 and after['value']==23 and before['C24']==12 and before['C25']==10 and after['C24']==9 and after['C25']==6
    assert before['font_bold'] and before['font_size']==13 and before['number_format']=='0' and before['column_width']==55
    for k in set(before)-{'value','C24','C25'}:assert before[k]==after[k],k
    ignored={(native.sheet_part(H.reg(r)),'C24'),(native.sheet_part(H.reg(r)),'C25')}
    backup=Path(c['operation']['backup']['path'])
    assert native.package(backup,ignored)==native.package(H.reg(r),ignored)
    streams=native.vba_streams(backup);assert len(streams)==11 and streams==native.vba_streams(H.reg(r))
    return {'sandbox':str(r),'before':before,'after':after,'bindings':[bb,ab],'all_11_VBA_streams':streams,'all_package_parts_compared':True,'normal_open_CorruptLoad_0':True}
if __name__=='__main__':
    with office.diagnostic_scope(OUT/'office'):
        case('private_DACL_source_copy_external_races',permissions_and_lock)
        for n in (259,260,261,370):case('native_source_length_'+str(n),lambda n=n:long_native(n))
        case('copy_changed_bytes_before_native_STOP',content_before)
        case('copy_changed_identity_cleanup_preserves_foreign',identity_before)
        for k in ('copy_sha','source_sha','copy_identity','source_identity'):case('post_native_'+k+'_STOP',lambda k=k:after_check(k))
        case('native_exception_cleanup_primary_preserved',ordinary_exception)
        case('workspace_uuid_collision_STOP_no_reuse_no_intent',directory_collision)
        case('exclusive_copy_collision_foreign_preserved',collision_foreign)
        case('wrong_guard_STOP',guard_wrong)
        for k in ('missing','file','relative','long','native_failure'):case('close_preflight_'+k,lambda k=k:no_preflight_mutation(k))
        case('close_preflight_real_denied_ACL',acl_denied)
        case('restore_preflight_no_rename',lambda:restore_preflight(False))
        case('resume_preflight_no_rename',lambda:restore_preflight(True))
        case('two_concurrent_close_once',H.parallel)
        case('existing_writer_STOP',lambda:H.preexisting_writer(False))
        case('existing_writable_mapping_STOP',lambda:H.preexisting_writer(True))
        case('same_SHA_unknown_identity_STOP',H.identity_only)
        case('native_styles_links_visibility_and_11_streams',independent_native_styles)
        case('full_native_package_and_renamed_request_idempotency',H.positive)
        case('RPC_E_DISCONNECTED_Save_exact_safe_STOP',lambda:H.disconnect(True))
        case('real_owned_Excel_Quit_Save_disconnect_safe_STOP',lambda:H.disconnect(False))
        case('Workbook_Close_primary_hresult_retained',H.book_close_primary)
        case('torn_native_Save_private_copy_safe_STOP',H.torn_stage)
        case('no_UZK_no_registry_write',H.no_uzk)
    print('DONE',sum(x['status']=='PASS' for x in RESULTS),sum(x['status']=='FAIL' for x in RESULTS),flush=True)
