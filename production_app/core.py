from pathlib import Path
from dataclasses import replace
from datetime import date,datetime,timezone
from uuid import uuid4
from decimal import Decimal
import hashlib,os,shutil
from copy import copy
from passport_app.registry import read_registry,append_row
from passport_app.excel import excel_session,open_excel,locked_file,atomic_replace,load_book
from . import model,facts as factual,office
from .storage import Stop,safe,contained,sha,digest,save,read,local_root,pins,verify_pins,mutex
from stage8_closure.storage import frozen

def journal_write(path,state):save(path,state|{'seal':digest(state)})

def journal_read(path):
    d=read(path);seal=d.pop('seal',None)
    if digest(d)!=seal:raise Stop('Журнал повреждён, автоматическая запись/восстановление запрещены: '+str(path))
    return d

def training_copy(cfg,parent):
    """A dedicated, explicitly chosen copy. Never infer production destinations."""
    source,documents,output=model.validate_paths(cfg)
    root=safe(parent)/('Учебный_'+uuid4().hex[:8]);root.mkdir(exist_ok=False)
    target=root/'COPY_ВЗМК.xlsx'
    with locked_file(source) as payload:target.write_bytes(payload)
    mark={'schema':'VZMK_TRAINING_COPY_1','registry':str(target),'source':str(source),'original_sha256':sha(source),'id':root.name}
    save(root/'TRAINING_COPY.json',mark)
    # Preserve the separately selected historical source root READ ONLY.
    # Before training, `output` is a user-selected folder containing the
    # original passports; after creating the sandbox, output points to TEMP.
    archive=cfg.get('first_passport_documents') or cfg['output']
    return cfg|{'registry':str(target),'output':str(root),
                'first_passport_documents':str(archive),'training':True}

def training_guard(cfg):
    if not cfg.get('training'):return
    p=safe(cfg['registry']);mark=read(p.parent/'TRAINING_COPY.json')
    if p.name!='COPY_ВЗМК.xlsx' or mark.get('registry')!=str(p) or safe(mark['source'])==p:raise Stop('Учебный режим разрешён только на отдельной созданной программой копии реестра.')
    if safe(cfg['output'])!=p.parent:raise Stop('Учебные документы должны оставаться в папке учебной копии.')

def pending(registry):
    root=safe(registry).parent/'.vzmk-operations'
    return [p for p in root.glob('*/journal.json') if journal_read(p).get('phase') not in ('COMPLETE','ROLLED_BACK','ABORTED')]

def plan(cfg,project,number,vom,facts,new):
    registry,documents,output=model.validate_paths(cfg);training_guard(cfg)
    if pending(registry):raise Stop('Есть незавершённая операция. Откройте «Восстановление»: '+str(pending(registry)[0]))
    r=read_registry(registry);p=model.select(r,project);profile=model.profile(cfg,project)
    if new and number!=p.next_number:raise Stop('Следующий номер изменился: '+p.next_number)
    existing=[e for e in r.projects[project] if e['number']==number]
    if new and existing or not new and len(existing)!=1:raise Stop('Дубль/не найден паспорт: '+number)
    v=model.visible_vom(safe(vom))
    shared_uzk=None
    if cfg.get('use_shared_uzk'):
        from .shared_uzk import (
            find_shared_uzk,planned_control,verify_planned_quantities,
        )
        shared_uzk=find_shared_uzk(cfg)
        expected=planned_control(shared_uzk,project,v)
        verify_planned_quantities(facts,expected)
    if not new and abs(v.mass-Decimal(str(existing[0]['mass'])))>Decimal('.005'):
        raise Stop('Масса ВОМ не совпадает с существующим паспортом.')
    if facts.get('auto_registry_dates'):
        from .registry_dates import resolve_registry_dates
        # For a new row, today's ending date is only a PROPOSAL. The other
        # dates must result from valid source formulas on a private registry
        # copy, never copied historic literals. Show the day at confirmation.
        resolved=resolve_registry_dates(cfg,project,number,v,new)
        facts={**facts,**resolved}
    variants=factual.validate(facts,v,project,number,cfg.get('training',False))
    for n in variants:
        for role in (f'akz{n}',f'akz_appendix{n}'):
            if role not in profile['templates']:raise Stop('Нет собственного шаблона '+role)
    from .templates import validate as validate_templates
    validate_templates(profile,p,variants,cfg.get('training',False))
    # The historically detailed UZK form requires each real welded-joint row.
    from .word_controls import inspect_word_fields
    from .real_forms import validate_uzk_details
    uzk_tokens,_=inspect_word_fields(profile['templates']['uzk']['path'])
    validate_uzk_details(facts,v,required='UZK_DETAILS_TABLE' in uzk_tokens)
    from .real_forms import require_form_inputs
    require_form_inputs(profile,facts)
    from .real_forms import require_akz_form_inputs
    require_akz_form_inputs(profile,facts,variants)
    # Pin the exact FIRST passport source files alongside inputs. Even the
    # historical acts used for visual provenance (not generation) must not
    # silently change between CHECK and PREPARE. Dedupe pinned paths.
    first_refs=profile.get('first_passport_sources',{}).get('sources',{})
    inputs=list(dict.fromkeys([
        registry,v.path,
        *[t['path'] for t in profile['templates'].values()],
        *[t['path'] for t in first_refs.values()],
        *[e['path'] for e in facts['evidence']],
    ]))
    catalog=profile.get('catalog')
    if catalog and Path(catalog).is_file():inputs.append(catalog)
    if shared_uzk is not None:inputs.append(shared_uzk)
    # Frozen human entries supply explicitly confirmed assignments when no catalog exists.
    target=output/project/'Паспорт'/(model.filename(number)+' от '+date.fromisoformat(facts['date']).strftime('%d.%m.%Y'))
    safe(target)
    if target.exists():raise Stop('Папка готового комплекта уже существует: '+str(target))
    hashes=pins(inputs)
    if hashes[str(registry)]!=r.digest or hashes[str(v.path)]!=v.digest:raise Stop('Источник изменился во время проверки.')
    model.profile(cfg,project)
    for path in inputs:
        with locked_file(path):pass
    data={'schema':'VZMK_PRODUCTION_PLAN_1','id':uuid4().hex,'cfg':cfg,'project':project,'number':number,
          'new':bool(new),'facts':facts,'vom':str(v.path),'fingerprints':hashes,'target':str(target),
          'marks':len(v.rows),'quantity':str(v.quantity),'mass_kg':str(v.mass),'area_m2':str(v.area),
          'variants':variants,'registry':str(registry),'date':facts['date']}
    run=local_root()/'jobs'/data['id'];run.mkdir(parents=True,exist_ok=False)
    data['run']=str(run);data['seal']=digest(data)
    save(run/'plan.json',data)
    return {'plan':str(run/'plan.json'),'data':data,'status':'PASS'}

def load_plan(path):
    d=read(path);seal=d.pop('seal')
    if digest(d)!=seal:raise Stop('План повреждён.')
    d['seal']=seal
    if safe(path)!=local_root()/'jobs'/d['id']/'plan.json':raise Stop('План вне частной рабочей папки.')
    verify_pins(d['fingerprints']);training_guard(d['cfg'])
    return d

def prepare(path):
    d=load_plan(path)
    with frozen(d['fingerprints']):
        verify_pins(d['fingerprints'])
        result=_prepare(d)
        verify_pins(d['fingerprints'])
        return result

def _prepare(d):
    run=safe(d['run']);outputs=run/'documents'
    outputs.mkdir(exist_ok=False)
    registry=Path(d['registry']);snapshot=run/'registry.xlsx'
    payload=registry.read_bytes()
    if hashlib.sha256(payload).hexdigest()!=d['fingerprints'][str(registry)]:raise Stop('SHA реестра изменился.')
    snapshot.write_bytes(payload)
    v=model.visible_vom(d['vom']);r=read_registry(snapshot);p=model.select(r,d['project'])
    if d['new']:
        with excel_session() as app:
            row,_=append_row(app,snapshot,r,p,v,date.fromisoformat(d['date']))
    else:row=next(e['row'] for e in r.projects[d['project']] if e['number']==d['number'])
    # Recalculate formulas only in the private registry.
    with excel_session() as app:
        b=open_excel(app,snapshot)
        try:
            if d['new']:
                col=model.columns(r)['START_DATE'];cell=b.Worksheets(r.sheet).Cells(row,col)
                if not cell.HasFormula:cell.Value2=(date.fromisoformat(d['facts']['start_date'])-date(1899,12,30)).days
            app.CalculateFullRebuild();b.Save()
        finally:b.Close(SaveChanges=False)
    staged=read_registry(snapshot)
    if d['new']:verify_registry(registry,snapshot,d['project'],d['number'],v,date.fromisoformat(d['date']))
    out=office.render(model.profile(d['cfg'],d['project']),p,staged,row,v,d['facts'],outputs,d['cfg'].get('training',False))
    save(run/'prepared.json',{'plan_seal':d['seal'],'outputs':out,'output_fingerprints':pins([o['path'] for o in out]),'registry':str(snapshot),'row':row,'snapshot_sha256':sha(snapshot)})
    return {'status':'PASS','folder':str(outputs),'files':len(out)}

def verify_registry(original,after,project,number,vom,expected_date=None):
    r0=read_registry(original);r1=read_registry(after)
    entries=[e for e in r1.projects[project] if e['number']==number]
    if len(entries)!=1 or sum(map(len,r1.projects.values()))!=sum(map(len,r0.projects.values()))+1:raise Stop('Реестр должен содержать ровно одну новую строку.')
    e=entries[0]
    if abs(Decimal(str(e['mass']))-vom.mass)>Decimal('.000001') or e['connections']!=vom.connections:raise Stop('Неверные масса/соединения новой строки.')
    if expected_date is not None:
        actual=e['date'].date() if isinstance(e['date'],datetime) else e['date']
        if actual!=expected_date:raise Stop('Дата новой строки не совпала с подтверждённой датой паспорта.')
    a=load_book(original);b=load_book(after)
    try:
        if a.sheetnames!=b.sheetnames:raise Stop('Изменилась структура листов реестра.')
        for sa in a.worksheets:
            sb=b[sa.title]
            def print_ranges(sheet,extend=False):
                from openpyxl.worksheet.print_settings import PrintArea
                ranges=[]
                for area in PrintArea.from_string(str(sheet.print_area)).ranges if sheet.print_area else []:
                    bounds=list(area.bounds)
                    # Excel extends an existing print area ending at the last
                    # table row when ListRows.Add inserts the single new row.
                    if extend and sa.title==r0.sheet and bounds[3]==e['row']-1:bounds[3]+=1
                    ranges.append(tuple(bounds))
                return sorted(ranges)
            if print_ranges(sa,True)!=print_ranges(sb) or sa.print_title_rows!=sb.print_title_rows or sa.print_title_cols!=sb.print_title_cols:
                raise Stop('Изменилась область печати/повторяемые строки реестра: '+sa.title)
            for row in sa:
                for ca in row:
                    # Insertion must not change any pre-existing cell's value/formula/style.
                    cb=sb.cell(ca.row,ca.column)
                    if ca.value!=cb.value or ca.number_format!=cb.number_format or copy(ca.font)!=copy(cb.font) or copy(ca.border)!=copy(cb.border) or copy(ca.alignment)!=copy(cb.alignment) or copy(ca.fill)!=copy(cb.fill) or copy(ca.protection)!=copy(cb.protection):
                        raise Stop('Изменилась исходная ячейка реестра: '+sa.title+'!'+ca.coordinate)
        source=model.select(r0,project)
        s0=a[r0.sheet];s1=b[r1.sheet]
        for col in range(1,r0.width+1):
            c=s0.cell(source.first_row,col);n=s1.cell(e['row'],col)
            if c.data_type=='f' and n.data_type!='f':raise Stop('Утрачена формула новой строки.')
        if len(b[r1.sheet].tables)!=len(a[r0.sheet].tables):raise Stop('Утрачена Excel-таблица.')
        from openpyxl.utils.cell import range_boundaries
        for sa in a.worksheets:
            sb=b[sa.title]
            if set(sa.tables)!=set(sb.tables):raise Stop('Изменён набор Excel-таблиц.')
            for name,table in sa.tables.items():
                old=sa.tables[name];new=sb.tables[name]
                bounds=list(range_boundaries(old.ref))
                if sa.title==r0.sheet and name==r0.table:bounds[3]+=1
                if tuple(bounds)!=range_boundaries(new.ref) or old.tableStyleInfo!=new.tableStyleInfo:
                    raise Stop('Неверное расширение/оформление Excel-таблицы: '+name)
    finally:a.close();b.close()
    return e['row']

def verify(path,preview=False):
    d=load_plan(path);run=Path(d['run']);prepared=read(run/'prepared.json')
    if prepared['plan_seal']!=d['seal'] or sha(prepared['registry'])!=prepared['snapshot_sha256']:
        raise Stop('Изменена подготовленная копия реестра.')
    if not prepared.get('output_fingerprints'):
        raise Stop('Комплект подготовлен предыдущей версией. Повторите подготовку новой заявкой.')
    verify_pins(prepared['output_fingerprints'])
    report=office.verify_documents(prepared['outputs'],d['number'],d['cfg'].get('training',False))
    hashes=pins([o['path'] for o in prepared['outputs']])
    preview_path=None
    if preview:
        from .print_pdf import create_print_pdf
        draft=run/'preview'
        draft.mkdir(exist_ok=False)
        copied=[]
        for item in prepared['outputs']:
            source=Path(item['path'])
            target=draft/source.name
            if target.exists():
                raise Stop('Дублирующийся документ чернового PDF: '+target.name)
            shutil.copyfile(source,target)
            copied.append({'role':item['role'],'path':str(target)})
        # The draft already has the 'ЧЕРНОВИК/УЧЕБНЫЙ' watermark; PDF
        # remains in PRIVATE job storage and is never published as issued.
        printout=create_print_pdf(copied,draft,Path(d['target']).name)
        preview_path=printout['output']['path']
        hashes.update(pins([preview_path]))
    save(run/'verified.json',{
        'plan_seal':d['seal'],'documents':report,'fingerprints':hashes,
        'preview_pdf':preview_path,
    })
    return {'status':'PASS','files':len(report),'documents':report,
            'preview_pdf':preview_path,'ready_to_confirm':True}

def confirm(path,responsible,reviewed):
    d=load_plan(path);run=Path(d['run']);verified=read(run/'verified.json')
    if not reviewed or responsible!=d['facts']['responsible']:raise Stop('Подтвердите просмотр всех документов и ответственное лицо.')
    verify_pins(verified['fingerprints'])
    confirmation={'plan_seal':d['seal'],'verified_sha256':sha(run/'verified.json'),'responsible':responsible,'reviewed':True,'created_at':datetime.now(timezone.utc).isoformat()}
    save(run/'confirmation.json',confirmation)
    return {'status':'PASS','confirmation':str(run/'confirmation.json')}

def supervisor_proof(result_path,task,path):
    result=read(result_path)
    if result.get('status')!='PASS' or result.get('task')!=task or result.get('plan')!=str(path):raise Stop('Нет PASS завершённого Office-запуска '+task)
    sup_path=safe(result['office_supervisor']);summary=read(sup_path)
    request=read(sup_path.parent/'JOB_REQUEST.json')
    arguments=request['arguments']
    if '--task' not in arguments or arguments[arguments.index('--task')+1]!=task:raise Stop('Office-доказательство другой операции.')
    if summary.get('owned_orphans')!=0 or summary.get('ownership_errors') or summary.get('worker_returncode')!=0:raise Stop('Office не закрыт корректно.')
    if any(r.get('exit_code')!=0 or r.get('still_running') or r.get('forced_termination_attempted') for r in summary['instances']):raise Stop('Office: аварийное завершение/остаточный процесс.')
    return result

def release(path):
    d=load_plan(path);run=Path(d['run']);prepared=read(run/'prepared.json');verified=read(run/'verified.json');confirmation=read(run/'confirmation.json')
    if confirmation['plan_seal']!=d['seal'] or confirmation['verified_sha256']!=sha(run/'verified.json'):raise Stop('Нет подтверждения.')
    verify_pins(verified['fingerprints'])
    output=run/'release';output.mkdir(exist_ok=False)
    files=office.release_copies(prepared['outputs'],output,d['cfg'].get('training',False))
    office.verify_documents(files,d['number'],d['cfg'].get('training',False),draft=False)
    # Only the FINAL, unwatermarked release copies enter the printable PDF.
    # A failed native Office export stops before any registry modification.
    from .print_pdf import create_print_pdf, verify_package_print
    printed=create_print_pdf(files,output,Path(d['target']).name)
    files.append(printed['output'])
    result={'plan_seal':d['seal'],'outputs':files,
            'fingerprints':pins([o['path'] for o in files]),
            'print_pdf':printed['manifest']}
    verify_package_print(result,Path(d['target']).name)
    save(run/'release.json',result)
    return {'status':'PASS','plan':str(path),'task':'release',
            'print_pdf':printed['manifest']['filename'],
            'print_pdf_pages':printed['manifest']['pages']}

def commit(path,release_result):
    d=load_plan(path);run=Path(d['run']);registry=safe(d['registry']);target=safe(d['target'])
    prepared=read(run/'prepared.json');verified=read(run/'verified.json');confirmation=read(run/'confirmation.json')
    supervisor_proof(release_result,'release',path)
    released=read(run/'release.json')
    if released['plan_seal']!=d['seal']:raise Stop('Другой выпуск.')
    verify_pins(released['fingerprints'])
    from .print_pdf import verify_package_print
    verify_package_print(released,target.name)
    if confirmation['plan_seal']!=d['seal'] or confirmation['verified_sha256']!=sha(run/'verified.json') or not confirmation['reviewed']:raise Stop('Нет действительного подтверждения выпуска.')
    verify_pins(verified['fingerprints']);v=model.visible_vom(d['vom'])
    factual.validate(d['facts']|{'documents_reviewed':True},v,d['project'],d['number'],d['cfg'].get('training',False),True)
    with mutex(registry),frozen(p for p in d['fingerprints'] if p!=str(registry)):
        if pending(registry):raise Stop('Незавершённая операция: используйте восстановление.')
        if target.exists():raise Stop('Папка уже существует, перезапись запрещена.')
        token=d['id'];stage=registry.parent/'.vzmk-operations'/token
        stage.mkdir(parents=True,exist_ok=False)
        backup=stage/'backup.xlsx';candidate=stage/'candidate.xlsx';journal=stage/'journal.json'
        state={'id':token,'phase':'STARTED','registry':str(registry),'target':str(target),
            'backup':str(backup),'original_sha256':d['fingerprints'][str(registry)],'new':d['new'],'plan':str(path)}
        journal_write(journal,state)
        try:
            with locked_file(registry,writable=d['new']) as original:
                if hashlib.sha256(original).hexdigest()!=state['original_sha256']:raise Stop('SHA реестра изменился перед выпуском.')
                verify_pins({p:h for p,h in d['fingerprints'].items() if p!=str(registry)})
                with backup.open('xb') as f:f.write(original);f.flush();os.fsync(f.fileno())
                if sha(backup)!=state['original_sha256']:raise Stop('SHA резервной копии не совпал.')
                state['phase']='BACKUP_VERIFIED';journal_write(journal,state)
                target.parent.mkdir(parents=True,exist_ok=True)
                private=target.parent/('.pending-'+token);private.mkdir(exist_ok=False)
                out=[]
                for o in released['outputs']:
                    destination=private/Path(o['path']).name;shutil.copyfile(o['path'],destination)
                    if sha(destination)!=released['fingerprints'][o['path']]:raise Stop('Изменилась копия документа.')
                    out.append(o|{'path':str(destination)})
                state['output_hashes']={Path(o['path']).name:sha(o['path']) for o in out}
                state['pending']=str(private)
                if d['new']:
                    shutil.copyfile(prepared['registry'],candidate)
                    if sha(candidate)!=prepared['snapshot_sha256']:raise Stop('Копия реестра изменилась.')
                    verify_registry(backup,candidate,d['project'],d['number'],v,date.fromisoformat(d['date']))
                    state['candidate_sha256']=sha(candidate)
                save(private/'ISSUE.json',{'project':d['project'],'number':d['number'],'training':bool(d['cfg'].get('training')),
                    'facts':d['facts'],'files':state['output_hashes'],'source_registry_sha256':state['original_sha256']})
                state['issue_sha256']=sha(private/'ISSUE.json')
                state['phase']='PUBLISH_INTENT';journal_write(journal,state)
                private.rename(target)
                state['phase']='DOCUMENTS_PUBLISHED';journal_write(journal,state)
                if d['new']:
                    state['phase']='REPLACE_INTENT';journal_write(journal,state)
                    atomic_replace(candidate,registry,backup=stage/'replaced_original.xlsx')
                    state['phase']='REGISTRY_REPLACED';journal_write(journal,state)
            if d['new']:verify_registry(backup,registry,d['project'],d['number'],v,date.fromisoformat(d['date']))
            elif sha(registry)!=state['original_sha256']:raise Stop('Изменился реестр существующего паспорта.')
            state['phase']='NEEDS_READBACK';journal_write(journal,state)
            return {'status':'PASS','state':'ТРЕБУЕТСЯ ПРОВЕРКА OFFICE','folder':str(target),'backup':str(backup),'journal':str(journal),'training':bool(d['cfg'].get('training'))}
        except BaseException as exc:
            state['error']=str(exc)
            if state['phase']=='STARTED':state['phase']='ABORTED'
            # Keep intent: a crash at ANY boundary is recoverable by hash inspection.
            journal_write(journal,state)
            raise Stop('Выпуск не завершён. Используйте восстановление: '+str(journal)+'; '+str(exc)) from exc

def readback(journal):
    s=journal_read(journal);d=read(s['plan']);registry=safe(s['registry']);target=safe(s['target'])
    if s['phase']!='NEEDS_READBACK':raise Stop('Неверная фаза проверки.')
    if sha(registry)!=(s['candidate_sha256'] if s['new'] else s['original_sha256']):raise Stop('Реестр изменился после записи.')
    for name,h in s['output_hashes'].items():
        if sha(target/name)!=h:raise Stop('Изменён документ: '+name)
    with excel_session() as app:
        b=app.Workbooks.Open(str(registry),UpdateLinks=0,ReadOnly=True,AddToMru=False)
        try:
            r=read_registry(registry)
            found=[e for e in r.projects[d['project']] if e['number']==d['number']]
            if len(found)!=1 or b.Worksheets(r.sheet).Cells(found[0]['row'],r.columns['number']).Value2!=d['number']:raise Stop('Excel: новая строка не подтверждена.')
        finally:b.Close(SaveChanges=False)
    released=read(Path(s['plan']).parent/'release.json')
    from .print_pdf import verify_package_print,verify_print_pdf
    print_manifest=verify_package_print(released,target.name)
    name=print_manifest['filename']
    if name not in s['output_hashes'] or s['output_hashes'][name]!=print_manifest['sha256']:
        raise Stop('Сводный PDF не совпадает с манифестом выпуска.')
    verify_print_pdf(target/name,expected_pages=print_manifest['pages'])
    docs=[{'path':str(target/name),'role':name.rsplit(' - ',1)[-1].rsplit('.',1)[0]}
          for name in s['output_hashes'] if not name.lower().endswith('.pdf')]
    office.verify_documents(docs,d['number'],d['cfg'].get('training',False),False)
    return {'status':'PASS','task':'readback','plan':s['plan'],'journal':str(journal),
            'print_pdf_pages':print_manifest['pages']}

def complete(journal,readback_result):
    journal=safe(journal);s=journal_read(journal);registry=safe(s['registry'])
    supervisor_proof(readback_result,'readback',s['plan'])
    with mutex(registry),locked_file(registry) as current:
        if s['phase']!='NEEDS_READBACK':raise Stop('Неверная фаза завершения.')
        if hashlib.sha256(current).hexdigest()!=(s['candidate_sha256'] if s['new'] else s['original_sha256']):raise Stop('Реестр изменился после проверки.')
        for name,h in s['output_hashes'].items():
            if sha(safe(s['target'])/name)!=h:raise Stop('Документ изменился после проверки.')
        print_name=Path(s['target']).name+'.pdf'
        if print_name not in s['output_hashes']:
            raise Stop('Не найден проверенный сводный PDF. Завершение запрещено.')
        s['phase']='COMPLETE';journal_write(journal,s)
    return {'status':'PASS','state':'ГОТОВО','folder':s['target'],'journal':str(journal),
            'backup':s['backup'],'print_pdf':str(Path(s['target'])/print_name)}

def recover(journal):
    journal=safe(journal);s=journal_read(journal);registry=safe(s['registry'])
    if journal.parent.parent!=registry.parent/'.vzmk-operations' or journal.parent.name!=s['id']:raise Stop('Журнал вне папки операций выбранного реестра.')
    if s['phase'] in ('COMPLETE','ROLLED_BACK','ABORTED'):return {'status':'PASS','phase':s['phase']}
    backup=journal.parent/'backup.xlsx'
    if not backup.exists() and s['phase']=='STARTED':
        with mutex(registry),locked_file(registry,writable=True) as current:
            if hashlib.sha256(current).hexdigest()!=s['original_sha256'] or Path(s['target']).exists():raise Stop('Неизвестное состояние до резервного копирования.')
            s['phase']='ABORTED';journal_write(journal,s)
        return {'status':'PASS','phase':'ABORTED'}
    if safe(s['backup'])!=backup or not backup.is_file() or sha(backup)!=s['original_sha256']:raise Stop('Резервная копия отсутствует/повреждена. Ручная проверка; запись запрещена.')
    with mutex(registry),locked_file(registry,writable=True) as current:
        h=hashlib.sha256(current).hexdigest()
        if h not in (s['original_sha256'],s.get('candidate_sha256')):raise Stop('Реестр изменён другим пользователем. Откат запрещён.')
        target=safe(s['target'])
        if target.exists():
            if not s.get('output_hashes') or sha(target/'ISSUE.json')!=s.get('issue_sha256'):raise Stop('Комплект изменён после сбоя. Ручная проверка.')
            expected=set(s['output_hashes'])|{'ISSUE.json'}
            if {p.name for p in target.iterdir()}!=expected:raise Stop('В папке появились чужие файлы. Откат запрещён.')
            for name,expected_hash in s['output_hashes'].items():
                if sha(target/name)!=expected_hash:raise Stop('Готовый документ изменён. Откат запрещён.')
        if h!=s['original_sha256']:
            restore=journal.parent/('restore_'+uuid4().hex+'.xlsx');shutil.copyfile(backup,restore)
            atomic_replace(restore,registry,backup=journal.parent/('pre_restore_'+uuid4().hex+'.xlsx'))
        if target.exists():target.rename(journal.parent/'unpublished_documents')
        if h!=s['original_sha256'] and sha(registry)!=s['original_sha256']:raise Stop('Не совпал SHA восстановленного реестра.')
        s['phase']='ROLLED_BACK';journal_write(journal,s)
        return {'status':'PASS','phase':'ROLLED_BACK','journal':str(journal),'registry_sha256':s['original_sha256']}
