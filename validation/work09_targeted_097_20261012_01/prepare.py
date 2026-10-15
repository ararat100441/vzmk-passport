from audit import *
import shutil
tag='w097_'+__import__('uuid').uuid4().hex[:8]
base=Path(os.environ['LOCALAPPDATA'])/tag;parent=Path.home()/('V9W09_'+tag)
base.mkdir();parent.mkdir();temp=base/'tmp';temp.mkdir()
save(OUT/'SETUP.json',{'base':str(base),'temp':str(temp),'parent':str(parent),'source':str(Path(read(ROOT/'validation/stage9_092_20261012_01/DEV_SOURCE_COPY_INTEGRITY.json')['source'])),'vom':str(Path(read(ROOT/'validation/stage9_092_20261012_01/DEV_SOURCE_COPY_INTEGRITY.json')['source'])/'TEST_VOM.xls'),'exe_sha256':sha(EXE)})
adapts=[]
for name in ('e2e','lifecycle','normal_overlap','overlap_crash_cleanup','foreign_office','ownership_stop','negative','inventory','window_close_guard'):
    src=ROOT/'tests/stage9_092'/(name+'.py');s=src.read_text(encoding='utf8')
    if name=='lifecycle':
        start=s.index('        for pid,before in users.items():');end=s.index('        cases.append(',start)
        s=s[:start]+s[end:]
        s=s.replace("'user_office_preserved':True","'user_session_continuity':'OBSERVATIONAL_ONLY; controlled foreign test separately'")
    if name=='foreign_office':s=s.replace("str(ROOT/'tests/stage9_092/ownership_stop.py')","str(OUT/'ownership_stop.py')")
    if name in ('normal_overlap','overlap_crash_cleanup'):s=s.replace("SETUP['probe_workspace']","read(OUT/'LIFECYCLE_FINAL.json')['workspace']")
    if name=='negative':
        s=s.replace("old=read(next((OUT/'e2e_918edab0').glob('import_*.json')))['workspace']","long_parent=WORKSPACES/('long_'+('L'*95));long_parent.mkdir();old=task(RUN,'import',source=SOURCE,parent=long_parent)[0]['workspace']")
    if name=='inventory':s=s.replace("oldp=OUT/'source_before'/p.relative_to(ROOT)","oldp=ROOT/'validation/stage9_092_20261012_01/source_before'/p.relative_to(ROOT)")
    (OUT/(name+'.py')).write_text('# Adapted independent run controller; original SHA '+sha(src)+'\n'+s,encoding='utf8')
    adapts.append({'original':str(src),'original_sha256':sha(src),'controller':name+'.py','controller_sha256':sha(OUT/(name+'.py'))})
for name in ('ACCEPTED_CODE_BEFORE.json','BUILD_091_BEFORE.json'):
    shutil.copyfile(ROOT/'validation/stage9_092_20261012_01'/name,OUT/name)
save(OUT/'CONTROLLER_PROVENANCE.json',{'adapted_protocols':adapts,'independent_assertions':'own EXE binding before every task, commands, returncodes, isolated paths, expected force policy; controls rerun independently','product_changed':False})
print('Prepared independent roots',base,parent,flush=True)
