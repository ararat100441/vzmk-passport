"""Mutation and native pair tests restricted to independent review sandbox."""
from review import *
from lxml import etree as ET
import unittest, io
from contextlib import ExitStack
from openpyxl import load_workbook
from akz_stage6 import core as c
resume='--resume' in sys.argv
S=E/'sandbox';S.mkdir(exist_ok=resume)
F=S/'fixtures';F.mkdir(exist_ok=resume);O=S/'output';O.mkdir(exist_ok=resume)
results=load(E/'independent_tests.json')['results'] if resume else []
audit={'production_attempts':[],'network_attempts':[]}
def auditor(event,args):
    if event=='open' and isinstance(args[0],(str,bytes,os.PathLike)):
        path=Path(os.fsdecode(args[0]))
        if path.name.casefold() in ('взмк.xlsx','узк.xlsм','узк.xlsm'):
            audit['production_attempts'].append(str(path));raise AssertionError('PRODUCTION DENIED')
    if event in ('socket.connect','socket.getaddrinfo'):
        audit['network_attempts'].append(str(args));raise AssertionError('NETWORK DENIED')
sys.addaudithook(auditor)
def check(name,fn):
    if any(r['name']==name and r['status']=='PASS' for r in results):return
    try:
        detail=fn();results.append({'name':name,'status':'PASS','detail':detail})
        print('PASS',name,flush=True)
    except Exception as exc:
        results.append({'name':name,'status':'FAIL','error':str(exc),'traceback':traceback.format_exc()})
        print('FAIL',name,str(exc),flush=True)
    save('independent_tests.json',{'results':results,'audit':audit})
def yes(condition):assert condition
def stop(fn,needle):
    try:fn()
    except Exception as exc:
        assert needle in str(exc),(needle,type(exc).__name__,str(exc))
        return {'exception':type(exc).__name__,'message':str(exc)}
    raise AssertionError('EXPECTED STOP '+needle)
N='http://schemas.openxmlformats.org/spreadsheetml/2006/main'
def tag(n):return '{'+N+'}'+n
def rewrite(path,fn):
    with zipfile.ZipFile(path) as z:parts={n:z.read(n) for n in z.namelist()}
    sheet=ET.fromstring(parts['xl/worksheets/sheet1.xml']);fn(sheet)
    parts['xl/worksheets/sheet1.xml']=ET.tostring(sheet,encoding='utf-8',xml_declaration=True)
    with zipfile.ZipFile(path,'w',zipfile.ZIP_DEFLATED) as z:
        for n,value in parts.items():z.writestr(n,value)
def cell(root,coord,value):
    import re
    rn=int(re.search(r'\d+',coord)[0]);sd=root.find(tag('sheetData'))
    row=next((r for r in sd if int(r.get('r'))==rn),None)
    if row is None:row=ET.SubElement(sd,tag('row'),r=str(rn))
    old=next((v for v in row if v.get('r')==coord),None)
    if old is not None:row.remove(old)
    if value is None:return
    new=ET.SubElement(row,tag('c'),r=coord)
    if isinstance(value,str):
        new.set('t','inlineStr');ET.SubElement(ET.SubElement(new,tag('is')),tag('t')).text=value
    else:ET.SubElement(new,tag('v')).text=str(value)
def isolated_copy(source,name):
    p=F/name;shutil.copyfile(source,p);return p
pv=load(I/'preview.json')['payload'];original_fixture=Path(load(I/'test_assignments.json')['fixture'])
registry=isolated_copy(c.REGISTRY,'DEV_registry.xlsx')
act=isolated_copy(c.ACT,'ACT_template.xlsx');application=isolated_copy(c.APPLICATION,'APPLICATION_template.xlsx')
catalog=isolated_copy(original_fixture,'TEST_catalog.xlsx')
vom=isolated_copy(pv['vom']['path'],'real_selected_VOM.xls')
visibility_xlsx=isolated_copy(original_fixture.parent/'SYNTHETIC_visibility.xlsx','SYNTHETIC_visibility.xlsx')
visibility_xls=isolated_copy(original_fixture.parent/'SYNTHETIC_visibility.xls','SYNTHETIC_visibility.xls')
initial={str(p):sha(p) for p in (registry,act,application,catalog,vom,visibility_xlsx,visibility_xls)}
def make_boundary(count):
    p=isolated_copy(visibility_xlsx,f'SYNTHETIC_{count}_rows.xlsx')
    def mutate(root):
        sd=root.find(tag('sheetData'))
        for r in list(sd):
            if int(r.get('r'))>=4:sd.remove(r)
        for n in range(count):
            values=('Бп2-5',f'SYNTHETIC row {n+1:02}',1,1.23,1.23,.456,.456)
            for col,v in zip('ABCDEFG',values):cell(root,f'{col}{n+4}',v)
        af=root.find(tag('autoFilter'))
        if af is not None:root.remove(af)
        root.find(tag('dimension')).set('ref',f'A1:G{count+3}')
    rewrite(p,mutate);return p
with patch.multiple(c,REGISTRY=registry,ACT=act,APPLICATION=application,FIXTURES=F,OUTPUT=O,
                    PINS={act:sha(act),application:sha(application)}):
    check('independent_catalog_exact_copy_195_3_numeric_zero_192_blank',lambda:yes(
        len(c.catalog(catalog))==195 and sum(v['assignment']==0 for v in c.catalog(catalog).values())==3
        and sum(v['assignment'] is None for v in c.catalog(catalog).values())==192))
    p=c.preview('Гараж Казахстан','9/2026',vom,catalog)
    save('independent_preview.json',{'seal':p.seal,'payload':p.payload})
    check('native_reader_real_104_visible_397_hidden',lambda:yes(len(p.payload['vom']['rows'])==104 and len(p.payload['vom']['hidden_rows'])==397))
    check('native_reader_visibility_xlsx_filter_and_manual_hidden',lambda:yes(
        [r['row'] for r in c.visible_vom(visibility_xlsx)['rows']]==[4,5] and c.visible_vom(visibility_xlsx)['hidden_rows']==[6,7]))
    check('native_reader_visibility_xls_filter_and_manual_hidden',lambda:yes(
        [r['row'] for r in c.visible_vom(visibility_xls)['rows']]==[4,5] and c.visible_vom(visibility_xls)['hidden_rows']==[6,7]))
    check('preview_preserves_all_sources',lambda:yes(all(sha(path)==h for path,h in initial.items())))
    row=dict(p.payload['vom']['rows'][0]);row.update(row=4,mark='Бп2-5',name='SYNTHETIC',quantity=2,
            mass_unit=3.25,mass_total=6.5,area_unit=1.25,area_total=2.5)
    assignments={'Бп2-5':{'assignment':0,'row':41}}
    match=lambda rows,a=assignments:c.matching({'rows':rows},a)
    check('A_blank_not_numeric_zero',lambda:stop(lambda:match([row],{'Бп2-5':{'assignment':None,'row':41}}),'NO_AKZ_ZERO_ROWS'))
    for variant in (1,2,3):
        check(f'B_variant_{variant}_excluded_from_zero',lambda variant=variant:stop(lambda:match([row],{'Бп2-5':{'assignment':variant,'row':41}}),'NO_AKZ_ZERO_ROWS'))
    bad=isolated_copy(catalog,'conflicting_duplicate_catalog.xlsx')
    rewrite(bad,lambda r:[cell(r,'A197','Бп2-5'),cell(r,'B197','TEST conflicting duplicate'),cell(r,'C197',1)])
    check('C_duplicate_conflicting_mark_STOP',lambda:stop(lambda:c.catalog(bad),'CATALOG_DUPLICATE_CONFLICT'))
    check('D_missing_exact_raw_mark_STOP_when_no_selected_group',lambda:stop(lambda:match([{**row,'mark':'Бп2-5 '}]),'NO_AKZ_ZERO_ROWS'))
    for field in ('area_unit','area_total'):
        for value in (None,0,-1,'invalid'):
            check(f'E_{field}_{value!r}_STOP',lambda field=field,value=value:stop(lambda:match([{**row,field:value}]),'NUMBER_'))
    check('F_conflicting_unit_total_area_STOP',lambda:stop(lambda:match([{**row,'area_total':2.501}]),'AREA_CONFLICT_STOP'))
    check('mass_conflict_STOP',lambda:stop(lambda:match([{**row,'mass_total':6.51}]),'MASS_CONFLICT_STOP'))
    original_act=act.read_bytes();act.unlink()
    check('G_missing_template_STOP',lambda:stop(c.templates,'TEMPLATE_SHA_CHANGED_STOP'));act.write_bytes(original_act)
    act.write_bytes(original_act+b'ISOLATED MUTATION')
    check('G_changed_template_STOP',lambda:stop(c.templates,'TEMPLATE_SHA_CHANGED_STOP'));act.write_bytes(original_act)
    name='existing';target=O/name;target.mkdir(exist_ok=resume);(target/'sentinel.txt').write_text('DO NOT OVERWRITE')
    check('H_existing_output_no_overwrite',lambda:stop(lambda:c.generate(p,name),'OUTPUT_ALREADY_EXISTS_STOP'))
    check('H_existing_output_preserved',lambda:yes((target/'sentinel.txt').read_text()=='DO NOT OVERWRITE' and len(list(target.iterdir()))==1))
    for path in (registry,act,application,catalog,vom):
        old=path.read_bytes();path.write_bytes(old+b'ISOLATED AFTER PREVIEW MUTATION')
        check('I_changed_after_preview_'+path.name,lambda:stop(lambda:c.generate(p,'changed'),'INPUT_CHANGED_AFTER_PREVIEW_STOP'))
        path.write_bytes(old)
    check('preview_payload_changed_STOP',lambda:stop(lambda:c.verify_preview(c.Preview({**p.payload,'date':'2026-10-06'},p.seal)),'PREVIEW_CHANGED_STOP'))
    v17=make_boundary(17);v18=make_boundary(18)
    p17=c.preview('Гараж Казахстан','9/2026',v17,catalog,synthetic=True)
    check('17_rows_all_selected_and_ordered',lambda:yes(len(p17.payload['selected'])==17 and
        [r['row'] for r in p17.payload['selected']]==list(range(4,21)) and
        p17.payload['totals']=={'quantity':Decimal(17),'mass_total':Decimal('20.91'),'area_total':Decimal('7.752')}))
    check('J_18_rows_preview_STOP_no_truncation',lambda:stop(lambda:c.preview('Гараж Казахстан','9/2026',v18,catalog,synthetic=True),'TEMPLATE_CAPACITY_STOP'))
    check('K_no_variant_zero_no_fabricated_positive',lambda:stop(lambda:match([row],{}),'NO_AKZ_ZERO_ROWS'))
    all_blank=isolated_copy(catalog,'TEST_all_blank.xlsx')
    rewrite(all_blank,lambda r:[cell(r,f'C{n}',None) for n in (41,47,48)])
    check('K_actual_blank_catalog_preview_STOP',lambda:stop(lambda:c.preview('Гараж Казахстан','9/2026',vom,all_blank),'NO_AKZ_ZERO_ROWS'))
    mixed=isolated_copy(catalog,'TEST_mixed_variants.xlsx')
    rewrite(mixed,lambda r:[cell(r,'C41',1),cell(r,'C47',2)])
    check('B_actual_mixed_variants_selects_only_zero',lambda:yes(
        [r['mark'] for r in c.preview('Гараж Казахстан','9/2026',vom,mixed).payload['selected']]==['Бп2-19']))
    unsupported=isolated_copy(catalog,'TEST_unsupported_variant.xlsx')
    rewrite(unsupported,lambda r:cell(r,'C41',4))
    check('unsupported_variant_no_fallback',lambda:stop(lambda:c.catalog(unsupported),'AKZ_ASSIGNMENT_INVALID'))
    missing_area=isolated_copy(v17,'SYNTHETIC_missing_area.xlsx')
    rewrite(missing_area,lambda r:cell(r,'F4',None))
    check('E_native_missing_area_preview_STOP',lambda:stop(lambda:c.preview('Гараж Казахстан','9/2026',missing_area,catalog,synthetic=True),'NUMBER_MISSING'))
    conflicting_area=isolated_copy(v17,'SYNTHETIC_conflicting_area.xlsx')
    rewrite(conflicting_area,lambda r:cell(r,'G4',.457))
    check('F_native_conflicting_area_preview_STOP',lambda:stop(lambda:c.preview('Гараж Казахстан','9/2026',conflicting_area,catalog,synthetic=True),'AREA_CONFLICT_STOP'))
    def second_failure(*args):raise c.Stage6Error('INDEPENDENT_SECOND_BOOK_FAILURE')
    with patch.object(c,'write_application',side_effect=second_failure):
        check('second_book_failure_STOP',lambda:stop(lambda:c.generate(p17,'second_failure'),'INDEPENDENT_SECOND_BOOK_FAILURE'))
    check('second_book_failure_no_published_or_pending_pair',lambda:yes(not (O/'second_failure').exists() and not list(O.glob('.pending-*'))))
    with patch.object(c,'validate_saved_pair',side_effect=c.Stage6Error('INDEPENDENT_SAVED_VALIDATION_FAILURE')):
        check('saved_validation_failure_STOP',lambda:stop(lambda:c.generate(p17,'validation_failure'),'INDEPENDENT_SAVED_VALIDATION_FAILURE'))
    check('saved_validation_failure_no_partial_pair',lambda:yes(not (O/'validation_failure').exists() and not list(O.glob('.pending-*'))))
    import win32file,win32con
    original_open=c.native_open
    def locked_second_open(app,path,*,read_only):
        if Path(path).name=='ТЕСТ Приложение к АКЗ №0.xlsx' and not read_only:
            handle=win32file.CreateFile(str(path),win32con.GENERIC_READ,0,None,win32con.OPEN_EXISTING,0,None)
            try:return original_open(app,path,read_only=read_only)
            finally:handle.Close()
        return original_open(app,path,read_only=read_only)
    with patch.object(c,'native_open',side_effect=locked_second_open):
        check('locked_second_output_actionable_native_error',lambda:stop(lambda:c.generate(p17,'locked_pair'),'Microsoft Excel'))
    check('locked_output_no_partial_pair',lambda:yes(not (O/'locked_pair').exists() and not list(O.glob('.pending-*'))))
    old_registry=registry.read_bytes();original_write_application=c.write_application
    def changed_during_generation(*args):
        original_write_application(*args);registry.write_bytes(old_registry+b'ISOLATED MID GENERATION CHANGE')
    with patch.object(c,'write_application',side_effect=changed_during_generation):
        check('input_changed_during_generation_STOP',lambda:stop(lambda:c.generate(p17,'mid_change'),'INPUT_CHANGED_AFTER_PREVIEW_STOP'))
    registry.write_bytes(old_registry)
    check('input_changed_during_generation_no_partial_pair',lambda:yes(not (O/'mid_change').exists() and not list(O.glob('.pending-*'))))
    def verify_17_generated():
        output=c.generate(p17,'SYNTHETIC_SUPPORTED_17')
        book=load_workbook(output/'ТЕСТ Приложение к АКЗ №0.xlsx',data_only=True)
        try:
            s=book.active
            assert [s[f'B{r}'].value for r in range(4,21)]==[f'SYNTHETIC row {r:02}' for r in range(1,18)]
            assert all(s[f'A{r}'].value=='Бп2-5' for r in range(4,21))
            assert s['C21'].value==17 and abs(s['E21'].value-20.91)<1e-8 and abs(s['G21'].value-7.752)<1e-8
            assert load(output/'PAIR.json')['state']=='COMPLETE_TEST_PAIR'
            return {'output':str(output),'row_count':17,'totals':[s[x].value for x in ('C21','E21','G21')]}
        finally:book.close()
    check('17_rows_native_generation_complete_no_loss',verify_17_generated)
    check('official_mode_STOP_before_access',lambda:stop(lambda:c.generate(None,'official',mode='OFFICIAL'),'OFFICIAL_MODE_STOP'))
    check('generate_without_preview_STOP',lambda:stop(lambda:c.generate(None,'no_preview'),'PREVIEW_CHANGED_STOP'))
    def gui_checks():
        import tkinter as tk
        from akz_stage6.gui import Launcher
        root=tk.Tk();root.withdraw()
        try:
            g=Launcher(root)
            assert str(g.generate_button['state'])=='disabled'
            with patch('akz_stage6.gui.messagebox.showerror') as error,patch.object(c,'generate') as generate:
                g.generate();assert not generate.called and 'PREVIEW_REQUIRED' in str(error.call_args)
            for var,new in ((g.project,'Гараж Казахстан'),(g.passport,'9/2026'),(g.vom,str(v17)),(g.catalog,str(catalog)),(g.synthetic,True)):
                with patch.object(c,'preview',return_value=p17):g.check()
                assert str(g.generate_button['state'])=='normal'
                var.set(new);root.update()
                assert g.checked is None and str(g.generate_button['state'])=='disabled'
            return {'real_tk_widget_checks':5,'no_generation_before_preview':True}
        finally:root.destroy()
    check('real_Tk_preview_required_and_all_selections_invalidate',gui_checks)
    suite=unittest.defaultTestLoader.discover(str(ROOT/'tests'),pattern='test_stage6.py')
    log=io.StringIO();unit=unittest.TextTestRunner(stream=log,verbosity=2).run(suite)
    (E/'unit_tests_independent.txt').write_text(log.getvalue(),encoding='utf-8')
    check('unchanged_product_unit_suite_16',lambda:yes(unit.wasSuccessful() and unit.testsRun==16))
    check('initial_sandbox_source_copies_restored',lambda:yes(all(sha(path)==h for path,h in initial.items())))
save('independent_tests.json',{'results':results,'audit':audit,'summary':{
    'PASS':sum(r['status']=='PASS' for r in results),'FAIL':sum(r['status']=='FAIL' for r in results)}})
print('SUMMARY',len(results),'FAIL',sum(r['status']=='FAIL' for r in results),flush=True)
