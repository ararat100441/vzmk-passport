"""Read-only artifact audit, report, and final evidence seal after frozen CLI tests."""
import difflib,hashlib,json,os,shutil,sys,time
from pathlib import Path
import win32api,win32con
OUT=Path(__file__).resolve().parent;ROOT=OUT.parents[1];PREV=ROOT/'validation/akz_settings_0911_20261013_01';REPRO=ROOT/'validation/akz_settings_0911_resume_20261013_01'
sys.path.insert(0,str(ROOT))
from stage8_closure import storage as S
from stage9_app import akz_settings as K
def read(p):return json.loads(Path(p).read_text(encoding='utf8'))
def sha(p):return hashlib.sha256(Path(p).read_bytes()).hexdigest()
def save(p,d):
    with Path(p).open('x',encoding='utf8') as f:json.dump(d,f,ensure_ascii=False,indent=2,default=str)
def check_map(d):
    for p,h in d.items():assert sha(p)==h,(p,'SHA_CHANGED')
def seal_ok(d):assert d['seal']==S.digest({k:v for k,v in d.items() if k!='seal'})
assert not (OUT/'FOREIGN_OFFICE_STOP.json').exists()
cases=read(OUT/'CLI_CASES.json');assert cases['status']=='PASS' and cases['new_cases']==8
office=read(OUT/'OFFICE_FINAL.json');assert office['status']=='PASS' and not office['foreign_stop']
legacy=read(OUT/'LEGACY_FINAL.json');ab=read(PREV/'AB_FINAL.json');build=read(OUT/'BUILD_MANIFEST.json');exe=Path(build['exe'])
check_map(read(OUT/'SOURCE_AT_BUILD.json'));check_map(read(PREV/'ACCEPTED_BEFORE.json'))
check_map(legacy['old_documents_manifests_receipts_unchanged']);check_map(legacy['old_day_files_unchanged'])
for key in ('A_files','B_files'):check_map(ab[key])
assert {str(p.relative_to(exe.parent)):sha(p) for p in exe.parent.rglob('*') if p.is_file()}==build['files']
before=read(OUT/'IMMUTABLE_BEFORE.json');original_sha=before['live_settings_original_sha256']
audit=[]
for label,manifest in [('LEGACY',legacy['manifest']),('A',ab['A_manifest']),('B',ab['B_manifest'])]:
    p=Path(manifest);m=read(p);seal_ok(m);assert m['official_issue'] is False
    receipts=sorted(p.parent.rglob('PAIR.json'));assert len(receipts)==4
    ref=m.get('akz_settings_snapshot')
    if ref:
        cfg=K.verify_snapshot(p.parent.parent.parent,p.parent,ref)
        assert cfg['climate']['air_c']==([15,25] if label=='A' else [16,26])
        assert not any(Path(f).parts[-3:]==('preferences','akz_auto','settings.json') for f in m['inputs'])
        assert win32api.GetFileAttributes(ref['path'])&win32con.FILE_ATTRIBUTE_READONLY
    rows=[]
    for receipt in receipts:
        d=read(receipt);assert d['official_issue'] is False and d['mode']=='TEST_SIMULATION'
        assert d['settings_sha256']==(ref['sha256'] if ref else original_sha)
        if ref:
            assert d['settings_snapshot']==ref
            assert m['inputs'][str(receipt)]==sha(receipt)
        for name,h in d['files'].items():assert sha(receipt.parent/name)==h
        rows.append({'path':str(receipt),'sha256':sha(receipt),'variant':d['variant'],'document_hashes_verified':True,'settings_binding_verified':True})
    assert sorted(r['variant'] for r in rows)==[0,1,2,3]
    audit.append({'label':label,'manifest':str(p),'sha256':sha(p),'seal_verified':True,'snapshot':ref,'receipts':rows,'rewritten':False})
reuse=next(j for j in cases['cases'] if j['label']=='LEGACY_FLOAT_DAY_REUSE');generated=read(reuse['result']);directory=Path(generated['pairs']['path'])
complete=read(directory/'COMPLETE.json');seal_ok(complete);ref=complete['settings_snapshot']
cfg=K.verify_snapshot(Path(legacy['manifest']).parent.parent.parent,Path(ref['path']).parent.parent,ref)
assert cfg['climate']['air_c']==[15,25] and len(complete['receipt_sha256'])==4
for relative,h in complete['receipt_sha256'].items():
    receipt=directory/relative;assert sha(receipt)==h;d=read(receipt)
    assert d['settings_snapshot']==ref and d['settings_sha256']==ref['sha256']
    for name,h in d['files'].items():assert sha(receipt.parent/name)==h
    assert d['validation']['status']=='PASS'
save(OUT/'ARTIFACT_READONLY_AUDIT.json',{'status':'PASS','kind':'READ_ONLY_PYTHON_AUDIT_NOT_EXE_OR_GUI','full_bundles':audit,'new_standalone_complete':str(directory/'COMPLETE.json'),'new_standalone_seal_and_4_receipts_verified':True,'legacy_climate_unchanged':True,'no_old_manifests_receipts_or_frozen_settings_written':True})
integrity=read(PREV/'SOURCE_INTEGRITY_FINAL.json')
for row in integrity['real_template_files']:assert sha(row['path'])==row['sha256']
seed=read(PREV/'SEED_BEFORE.json');check_map(seed['files'])
old_builds=[]
for base in [ROOT/'validation/akz_auto_20261013_02',PREV]:
    b=read(base/'BUILD_MANIFEST.json');folder=Path(b['exe']).parent
    assert {str(p.relative_to(folder)):sha(p) for p in folder.rglob('*') if p.is_file()}==b['files']
    old_builds.append({'exe':b['exe'],'sha256':b['exe_sha256'],'onedir_files':len(b['files']),'unchanged':True})
product={str(Path(p).relative_to(ROOT)):h for p,h in read(OUT/'SOURCE_AT_BUILD.json').items() if not Path(p).is_relative_to(ROOT/'tests')}
product_sha=hashlib.sha256(json.dumps(product,ensure_ascii=False,sort_keys=True,separators=(',',':')).encode('utf8')).hexdigest()
save(OUT/'SOURCE_INTEGRITY_FINAL.json',{'status':'PASS','accepted_stage1_8_unchanged':56,'seed_files_unchanged':len(seed['files']),'real_templates_unchanged':len(integrity['real_template_files']),'product_sources':product,'product_sha256':product_sha,'old_builds':old_builds,'changes_from_0911':read(OUT/'CHANGED_FILES.json'),'all_old_bundles_and_new_A_B_frozen_files_unchanged':True,'production_or_official_writes':False,'prior_wide_seal_cache_drift':read(OUT/'FINAL_BINDING.json')['prior_wide_seal_runtime_cache_drift']})
# Preserve the failed candidate's commands/results, including controller-only pre-dispatch failures.
for p in REPRO.rglob('*'):
    if p.is_file():
        target=OUT/'repro_0911'/p.relative_to(REPRO);target.parent.mkdir(parents=True,exist_ok=True);shutil.copyfile(p,target);assert sha(p)==sha(target)
changed=read(OUT/'CHANGED_FILES.json')
changes='\n'.join(f"- `{d['relative']}` — `{d['after_sha256']}`" for d in changed)
def link(name):return f'[{name}]({(OUT/name).as_posix()})'
report=f'''ВЗМК — завершение legacy-проверок и минимальный CRLF fix. Только DEV/TEST.

EXE 0.9.11: прежние21 frozen self-checks сохранены. Legacy проверка после смены диапазонов **FAIL**: LEGACY_AKZ_SETTINGS_UNPROVEN. Точный SHA0.9.11 `{old_builds[1]['sha256']}`; исходная onedir1120 файлов сохранена.

Дефект воспроизведён на EXE и отдельным source regression test до исправления. Windows workspace.save сохраняет JSON через text-mode с CRLF; legacy_evidence строила только LF bytes. В history есть настройки, чья CRLF сериализация побайтно совпадает с оригиналом, SHA `{original_sha}`; LF SHA отличается. {link('ROOT_CAUSE_REPRODUCTION.json')}. Это ошибка Stage9 восстановления evidence, не изменение диапазонов, формулы, нормативов или Stage1–8.

Исправление: проверить только два известных формата writer — LF и CRLF — и принять исключительно точное совпадение с SHA исходного sealed manifest. Численное сходство не считается доказательством. Старые файлы не переписываются. Добавлены Windows-writer regression и отрицательный тест численной эквивалентности без byte evidence. {link('EXACT_CODE_DIFF.patch')}, {link('CHANGED_FILES.json')}.

Кандидат **0.9.12**: `{exe}`. SHA-256 `{build['exe_sha256']}`. Product source+spec SHA `{product_sha}`. Onedir1120 файлов, 63 raw и52 compiled modules сверены с текущими исходниками. Зависимости и команда сборки: {link('BUILD_ENVIRONMENT.json')}, {link('BUILD_COMMAND.json')}; привязка модулей и полный перечень SHA: {link('BUILD_MANIFEST.json')}, {link('SOURCE_AT_BUILD.json')}. Сборки0.9.10/0.9.11 сохранены полностью.

**Legacy tests: PASS на0.9.12.** Выполнены8 новых frozen CLI cases, каждый в отдельном процессе EXE с PATH только Windows/System32. 7 результатов PASS и1 ожидаемый FAIL/STOP; все8 expectations PASS. {link('CLI_CASES.json')}:

- Эквивалентное сохранение float-диапазонов: no-op, точные bytes/SHA live settings прежние.
- Повторное формирование отдельного нового TEST-run четырёх АКЗ0/1/2/3: native Excel validation PASS; старые day-records повторно использованы, SHA неизменны. Прежний legacy-комплект не перегенерирован.
- Смена будущего диапазона15–25→16–26, затем проверка старого manifest: PASS, LEGACY_SETTINGS_RECOVERED_EXACT_SHA, CRLF, точный исходный SHA из history.
- Временная недоступность единственного подтверждающего history-файла в собственном TEST-стенде: ожидаемый FAIL LEGACY_AKZ_SETTINGS_UNPROVEN с предложением нового TEST-комплекта. Защита сохранена. History возвращён по прежнему пути с прежними bytes/SHA.
- Новый процесс EXE после восстановления evidence: legacy verify PASS.
- Существующие новые комплекты A15–25 и B16–26: оба verify PASS на0.9.12 при live settings B16–26. Каждый использует собственный frozen snapshot.

Настройки/manifests/receipts/docs старого legacy и новых A/B проверены без записи. Их исходные hashes, seals, frozen links и hashes документов неизменны. Для нового standalone-run проверены COMPLETE seal, snapshot и4 pinned receipts. {link('LEGACY_FINAL.json')}, {link('ARTIFACT_READONLY_AUDIT.json')}. Audit JSON структур — отдельная read-only Python проверка, а не результат GUI/EXE. После сценария восстановлены только первоначальные точные bytes изменяемого TEST live settings; старые manifests, receipts, документы, snapshots и климат не редактировались.

Свежие targeted source persistence tests **12/12 PASS** (включая2 новых). Ранее выполненные source42/42 и Office17/17 остаются историческими PASS0.9.11; повторный полный набор и Office17 на0.9.12 **NOT_RUN**: изменение затрагивает только legacy settings recovery, существующие Office/ownership модули byte-identical. Ранее native Word/Excel проверки полных A/B относятся к0.9.11; их выходные документы неизменны и повторно подтверждены SHA. На0.9.12 выполнена новая native Excel генерация четырёх вариантов; повторный полный Word/native A/B E2E **NOT_RUN** по targeted scope.

PID48680: birth2026-10-13T10:34:12.776Z, exit наблюдался10:53:08.102796Z, code0, причина **UNKNOWN**. PID отсутствует в списках своих Office. Application ±2мин:63 events, связанных EXCEL/PID matches0; Security вернул отсутствие matching events. Ни причина, ни закрывший актор не установлены; code0 не доказывает штатное закрытие пользователем. {link('PID48680_REVIEW.json')}; свежий event query в repro_0911/PID48680_EVENT_QUERY.json. Исторический Word3132 также UNKNOWN.

Office baseline перед0.9.12: Excel33924 и53328, retained read-only handles с PID+birth+image. Оба живы после тестов с теми же identities. Во время CLI ожидания controller проверял handles каждые0.1с. Чужие процессы не закрывались; новых foreign exits нет. Свои native Office4, force0, orphan0. {link('OFFICE_FINAL.json')}. Предыдущий controller проверял STOP после завершения синхронной команды; новый validation controller проверяет во время ожидания, код продукта/Office lifecycle не менялся.

Source integrity:56/56 принятых Stage1–8,32 sealed seed-файла,8 исходных образцов — неизменны. Реальные ВОМ/реестры/справочники/паспорта не записывались. {link('SOURCE_INTEGRITY_FINAL.json')}. Прежний широкий evidence seal дополнительно включал изменяемый Intel ShaderCache:1 cache SHA изменился после native запуска; это явно записано в FINAL_BINDING/SOURCE_INTEGRITY_FINAL, cache не восстановлен и исключён из нового immutable seal. Все защищённые source/input/run hashes проверены; старый evidence manifest и его741 файлов сохранены. Внешние3 source hashes изменены разрешённым Stage9 fix, их оригиналы сохранены и сверены. Новый seal не заявляет неизменность runtime cache.

Изменённые product/test/build файлы относительно0.9.11:

{changes}

Остальные изменения — текущие validation evidence/controllers и новая отдельная onedir. До EXE dispatch исправлена ошибка encoding в validation harness; ещё один pre-dispatch harness STOP выявил включённый в прежний seal runtime cache. Все STOP/FAIL сохранены в repro_0911 и текущем evidence; они не объявлены product PASS.

Оставшиеся продуктовые ошибки в проверенном targeted scope:0. Независимый Work review **NOT_RUN**; GUI Computer Use, GUI viewer/save/click E2E0.9.12 **NOT_RUN**, нового разрешения нет. Полный незатронутый Stage1–8 audit/production/официальный выпуск/deployment не выполнялись. Исторические Office exits остаются UNKNOWN.

READY FOR TARGETED WORK 09: **YES** — кандидат0.9.12 готов к независимому targeted review, сам review не запускался. READY FOR USER E2E: **NO**. READY FOR PRODUCTION: **NO**. {link('WORK09_ASSIGNMENT.md')}. После отчёта STOP.
'''
(OUT/'REPORT.md').write_text(report,encoding='utf8')
(OUT/'WORK09_ASSIGNMENT.md').write_text(f'''Targeted independent Work09 assignment — candidate0.9.12, DEV/TEST only. Do not dispatch Work from this assignment.

EXE `{exe}`; SHA `{build['exe_sha256']}`; product SHA `{product_sha}`. Verify EVIDENCE_MANIFEST_SHA256, full onedir1120, raw63/compiled52 and source binding first. Read REPORT, ROOT_CAUSE_REPRODUCTION, EXACT_CODE_DIFF, CHANGED_FILES, SOURCE_AT_BUILD, BUILD_MANIFEST, CLI_CASES, LEGACY_FINAL, ARTIFACT_READONLY_AUDIT, SOURCE_INTEGRITY_FINAL, OFFICE_FINAL and preserved0.9.11 failures in repro_0911. Previous baseline0.9.11 report remains a historical STOP; new report supersedes readiness, not history.

Review the two prior persistence MAJOR findings and this newly demonstrated CRLF regression, without a full Stage1–8 audit. Minimal0.9.12 change is LF/CRLF exact-SHA legacy reconstruction. Independently require exact bytes matching sealed manifest SHA. Equal numbers with different bytes, wrong/missing history, wrong settings, changed snapshot/receipt or same-date actual range conflict must STOP. Verify history is reconstructed in memory, old files never rewritten. On disposable copies repeat legacy equivalent float save, old climate reuse, future16–26 save, exact CRLF recovery, history-unavailable UNPROVEN, evidence restore and fresh EXE restart. Old manifest/docs/receipts/days SHA must remain identical. Check new A/B runs preserve independent frozen settings even after live changes/absence, with strict run/path/SHA/receipt seals; no arbitrary numeric fallback or protection bypass.

Review12 source persistence tests including Windows writer failure-before/fix-after and numeric-equivalence refusal. Prior source42/Office17 and21 EXE cases refer to0.9.11, not new0.9.12 results. New0.9.12 evidence is8 frozen CLI expectations (one expected FAIL) and4 native Excel instances. Native full Word/A/B on0.9.12 and GUI clicks are NOT_RUN; do not claim prior runtime tests were rerun. Retain Magnus, all variants/materials/layers/norms, АКЗ2 cumulative180 and original formatting/formulas/ВОМ semantics. No accepted algorithm changes.

Only isolated DEV/TEST copies; no real registry/VOM/catalog/template/passport changes, actual writeoff, official issuance or deployment. Before native tests capture fresh Office PID+creationtime+image and retain readonly handles. Foreign exit →immediate STOP, cause must remain UNKNOWN without evidence. Historical Excel48680 code0/Word3132 exits are unresolved; no attribution established. Never close foreign Office, global taskkill, unknown PID or request UAC. GUI Computer Use requires separate explicit permission. No source fix/build without independent demonstrated finding and user scope. Report independent verdict, exact proof/SHA, PASS/FAIL/NOT_RUN and remaining risks, then STOP.
''',encoding='utf8')
save(OUT/'RESULT.json',{'status':'PASS_TARGETED_SELF_VALIDATION','legacy_0911':'FAIL_REPRODUCED_CRLF','legacy_0912':'PASS','version':'0.9.12','exe':str(exe),'exe_sha256':build['exe_sha256'],'preserved_0911_sha256':old_builds[1]['sha256'],'product_sha256':product_sha,'fresh_source_persistence':12,'new_frozen_CLI_cases':8,'expected_FAIL_cases':1,'native_Excel_instances':office['own_office_instances'],'GUI':'NOT_RUN','independent_Work':'NOT_RUN','foreign_exit48680_cause':'UNKNOWN','READY_FOR_TARGETED_WORK_09':'YES','READY_FOR_USER_E2E':'NO','READY_FOR_PRODUCTION':'NO','report':str(OUT/'REPORT.md')})

def walk(folder):
    for base,dirs,names in os.walk(folder,followlinks=False):
        dirs[:]=[d for d in dirs if d!='__pycache__' and not (Path(base)/d).is_symlink() and not (Path(base)/d).is_junction()]
        for name in names:
            p=Path(base)/name
            if not p.is_symlink() and not p.is_junction():yield p
excluded={'EVIDENCE_MANIFEST.json','EVIDENCE_MANIFEST_SHA256.json'}
own={str(p.relative_to(OUT)):{'sha256':sha(p),'bytes':p.stat().st_size} for p in sorted(walk(OUT)) if p.name not in excluded}
external={}
def add(p):external[str(p)]={'sha256':sha(p),'bytes':Path(p).stat().st_size}
for folder in [Path(read(OUT/'SETUP.json')['legacy_workspace']),Path(read(OUT/'SETUP.json')['main_workspace']),Path(read(OUT/'SETUP.json')['base']),Path(read(REPRO/'SETUP.json')['base']),exe.parent,*[Path(b['exe']).parent for b in old_builds],PREV]:
    for p in walk(folder):
        if p.suffix.lower()!='.tmp':add(p)
for mapping in [read(OUT/'SOURCE_AT_BUILD.json'),read(PREV/'ACCEPTED_BEFORE.json'),seed['files']]:
    for p in mapping:add(Path(p))
for row in integrity['real_template_files']:add(Path(row['path']))
manifest={'schema':'STAGE9_TARGETED_LEGACY_RESUME_EVIDENCE_V1','state':'SEALED_SELF_VALIDATION','mode':'DEV_TEST_ONLY','official_issue':False,'version':'0.9.12','exe':str(exe),'exe_sha256':sha(exe),'product_sha256':product_sha,'at_utc':time.time(),'evidence_file_count':len(own),'evidence_files':own,'external_file_count':len(external),'external_files':external,'excluded':['manifest self and SHA sidecar','links and __pycache__','transient .tmp','driver/runtime cache outside protected TEST workspace and dedicated diagnostics'],'GUI':'NOT_RUN','independent_Work':'NOT_RUN','READY_FOR_TARGETED_WORK_09':'YES','READY_FOR_USER_E2E':'NO','READY_FOR_PRODUCTION':'NO','previous_sealed_stop_manifest':str(PREV/'EVIDENCE_MANIFEST.json'),'previous_manifest_sha256':sha(PREV/'EVIDENCE_MANIFEST.json')}
assert all(sha(OUT/p)==d['sha256'] for p,d in own.items());assert all(sha(p)==d['sha256'] for p,d in external.items())
save(OUT/'EVIDENCE_MANIFEST.json',manifest);seal=sha(OUT/'EVIDENCE_MANIFEST.json')
save(OUT/'EVIDENCE_MANIFEST_SHA256.json',{'manifest':str(OUT/'EVIDENCE_MANIFEST.json'),'sha256':seal,'status':'PASS','evidence_files':len(own),'external_files':len(external),'verification':'All listed SHA rechecked before and after publication; manifest self SHA verified'})
assert sha(OUT/'EVIDENCE_MANIFEST.json')==seal
assert all(sha(OUT/p)==d['sha256'] for p,d in own.items());assert all(sha(p)==d['sha256'] for p,d in external.items())
print('SEALED PASS',seal,'files',len(own),len(external),'READY_FOR_TARGETED_WORK_09 YES',flush=True)
