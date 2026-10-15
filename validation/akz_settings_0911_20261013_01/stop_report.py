from common import *
import difflib
stop=read(OUT/'FOREIGN_OFFICE_STOP.json');baseline=read(OUT/'OFFICE_BASELINE_FINAL.json');ab=read(OUT/'AB_FINAL.json');build=read(OUT/'BUILD_MANIFEST.json')
accepted=read(OUT/'ACCEPTED_BEFORE.json');assert all(sha(p)==h for p,h in accepted.items());sources=read(OUT/'SOURCE_AT_BUILD.json');assert all(sha(p)==h for p,h in sources.items())
old=read(ROOT/'validation/akz_auto_20261013_02/BUILD_MANIFEST.json');oldexe=Path(old['exe']);assert {str(p.relative_to(oldexe.parent)):sha(p) for p in oldexe.parent.rglob('*') if p.is_file()}==old['files']
assert {str(p.relative_to(EXE.parent)):sha(p) for p in EXE.parent.rglob('*') if p.is_file()}==build['files']
for field in ('A_files','B_files'):assert all(sha(p)==h for p,h in ab[field].items())
templates=read(ROOT/'validation/akz_auto_20261013_02/SOURCE_INTEGRITY_FINAL.json')['real_first_passport_template_pairs'];assert all(sha(p['path'])==p['sha256'] for p in templates)
seed=read(ROOT/'validation/akz_auto_20261013_02/SEED_BEFORE.json');assert all(sha(p)==h for p,h in seed['files'].items())
changes=[];diff=[]
for p in [ROOT/'stage9.py',*sorted((ROOT/'stage9_app').glob('*.py')),ROOT/'stage9_0911.spec',ROOT/'tests/stage9_akz_auto/test_persistence.py']:
    rel=p.relative_to(ROOT);before=OUT/'source_before'/rel;a=before.read_text(encoding='utf8') if before.exists() else '';b=p.read_text(encoding='utf8')
    if a==b:continue
    changes.append({'path':str(p),'relative':str(rel),'before_sha256':sha(before) if before.exists() else None,'after_sha256':sha(p)})
    diff+=difflib.unified_diff(a.splitlines(True),b.splitlines(True),fromfile='0.9.10/'+rel.as_posix(),tofile='0.9.11/'+rel.as_posix())
save(OUT/'CHANGED_FILES.json',changes);(OUT/'EXACT_CODE_DIFF.patch').write_text(''.join(diff),encoding='utf8')
product={str(Path(p).relative_to(ROOT)):h for p,h in sources.items() if not Path(p).is_relative_to(ROOT/'tests')};product_sha=hashlib.sha256(json.dumps(product,ensure_ascii=False,sort_keys=True,separators=(',',':')).encode('utf8')).hexdigest()
save(OUT/'SOURCE_INTEGRITY_FINAL.json',{'status':'PASS','accepted_stage1_8_files':accepted,'accepted_stage1_8_count':len(accepted),'immutable_seed_files':len(seed['files']),'real_template_files':templates,'preserved_0910_exe_sha256':sha(oldexe),'preserved_0910_onedir_files':len(old['files']),'new_0911_onedir_files':len(build['files']),'product_sources':product,'product_sha256':product_sha,'A_B_docs_manifests_receipts_snapshot_SHA_unchanged':True,'production_writes':False,'TEST_inputs_note':'Existing artificial assignment fixture copied byte-for-byte; accepted Stage1 added only TEST passport10/2026 to fresh TEST registry. No VOM/catalog content edits.'})
jobs=[];seen=set();foreign_pid=stop['identity']['ProcessId']
for p in sorted(OUT.glob('*.json')):
    d=read(p)
    if not isinstance(d,dict) or 'office_supervisor' not in d or d['office_supervisor'] in seen:continue
    seen.add(d['office_supervisor']);s=read(d['office_supervisor']);assert s['owned_orphans']==0 and not s['ownership_errors'];assert all(not i.get('forced_termination_attempted') and i['exit_code']==0 for i in s['instances'])
    assert all(i['pid']!=foreign_pid for i in s['instances'])
    cmd=read(p.with_suffix('.command.json'))['command'];version='0.9.11' if cmd[0]==str(EXE) else '0.9.10' if cmd[0]==SETUP['exe_0910'] else 'SOURCE'
    jobs.append({'result':str(p),'version':version,'status':d['status'],'supervisor':d['office_supervisor'],'office_instances':len(s['instances']),'own_pids':[i['pid'] for i in s['instances']]})
    folder=Path(d['office_supervisor']).parent
    assert folder.is_relative_to(Path(SETUP['base'])/'VZMKPassport/diagnostics')
    for file in folder.rglob('*'):
        if file.is_file():
            target=OUT/'office_jobs'/folder.name/file.relative_to(folder);target.parent.mkdir(parents=True,exist_ok=True);shutil.copyfile(file,target);assert sha(target)==sha(file)
events=json.loads((OUT/'FOREIGN_OFFICE_APPLICATION_EVENTS.json').read_text(encoding='utf-8-sig'))
save(OUT/'OFFICE_STOP_DIAGNOSIS.json',{'status':'STOP_FOREIGN_OFFICE_EXIT','foreign':stop,'baseline_final':baseline,'observed_at_utc':'2026-10-13T10:53:08.102796+00:00','relevant_Application_events':events,'cause':'UNKNOWN','attribution':'No causal attribution established; exit0 is not proof of who closed it','own_jobs':jobs,'foreign_pid_in_own_instance_lists':False,'owned_orphans':0,'own_forced_terminations':0,'monitor_limitation':'Flag written immediately; synchronous CLI controller inspected it after current job completed. That already-running worker was not interrupted mid-operation. No subsequent test was dispatched after the flag was consumed.','legacy_0911_checks':'NOT_RUN_STOP'})
def link(n):return f'[{n}]({(OUT/n).as_posix()})'
changed='\n'.join(f"- `{r['relative']}` — `{r['after_sha256']}`" for r in changes)
report=f'''ВЗМК — targeted AKZ settings persistence. Кандидат DEV EXE0.9.11. Итог: **STOP_FOREIGN_OFFICE_EXIT**.

**STOP и диагностика.** Сработало ранее заданное пользователем правило: «Если во время теста неожиданно завершится не принадлежащий тесту Office-процесс — немедленно STOP и диагностика». Исходный ExcelPID{foreign_pid}, creation time10:34:12.776 UTC, parent41204, завершился с exit_code0; наблюдение10:53:08.102796 UTC. Причина **UNKNOWN**, причинная связь с проверяемой программой не установлена. PID не входит ни в один список собственных Office instances в текущих supervisor logs; force termination0/owned orphan0. В ограниченном окне Application не найдено подходящих Office/ApplicationError/WER событий. Это не доказывает причину. PID33924 остаётся жив с исходной creation time.

Ограничение контроллера зафиксировано: монитор записал флаг сразу, но subprocess.run проверил его после завершения текущего CLI-задания. Текущий preserved0.9.10 worker не был прерван посередине и успел завершить своё формирование. Его process result PASS не принимается за PASS полного legacy-сценария. После получения флага следующие тесты не запускались; далее только чтение диагностики и запись отчёта/evidence. {link('FOREIGN_OFFICE_STOP.json')}, {link('OFFICE_STOP_DIAGNOSIS.json')}.

**1. ROOT CAUSE MAJOR1.** GUI0.9.10 преобразовывал все диапазоны float(); daily_climate вычислял S.digest сырого JSON. 15 и15.0 имели разные сериализации/digest. Численно неизменное сохранение вызывало AKZ_DAY_SETTINGS_CONFLICT. Независимое воспроизведение Work сохранено в `validation/work09_akz_auto_0910_20261013_01/SETTINGS_REPRESENTATION_AND_MANIFEST_DEFECT.json`.

**2. ROOT CAUSE MAJOR2.** _generate_existing копировал preview fingerprints, включая live preferences/akz_auto/settings.json, в manifest.inputs. verify_manifest сравнивал его старый SHA с текущими bytes. Поэтому будущая смена настроек инвалидировала неизменённый прежний комплект. Work evidence: MANIFEST_MUTABLE_SETTINGS_CONFIRMED.json.

**3. EXACT CODE DIFF.** {link('EXACT_CODE_DIFF.patch')}, {link('CHANGED_FILES.json')}. Изменения только Stage9, новый targeted test и отдельный build spec. Общий Stage8 S.digest, принятые алгоритмы, формулы ЛКМ, выбор ВОМ, Office ownership/supervisor не менялись.

**4. NUMERIC CANONICALIZATION.** Новый akz_settings.py формирует детерминированный JSON с сортировкой ключей и без лишних десятичных нулей/экспоненциального формата. 15/15.0/15.00 и−0/0 эквивалентны; реальные числовые различия сохраняются без округления. GUI использует Decimal, поддерживает запятую и отвергает roundtrip с потерей точности. Прежние допустимые диапазоны/точность остаются. Равное сохранение — no-op с прежними bytes/SHA/history; реальное изменение меняет digest. Старый климат проверяется по seal и точным int/float представлениям тех же шести границ (не более64), без переписывания/перегенерации. Новые day records содержат canonical schema/digest/settings. Реальная смена параметра на старой дате по-прежнему STOP.

**5. FROZEN SETTINGS IMPLEMENTATION.** В начале формирования нового комплекта создаётся exact-byte snapshot run/inputs/AKZ_SETTINGS.json с SHA-256/canonical digest. Файл read-only, создание поверх существующего запрещено. Manifest.inputs исключает live settings.json, включает snapshot и SHA receipts. Manifest/PAIR receipts связаны с одним snapshot. Проверяется path именно этого owned run, SHA/digest/schema, связь receipts и SHA документов; подмена/повреждение — STOP. Standalone auto-generation тоже создаёт run snapshot и sealed COMPLETE с SHA receipts. Для выключенного auto-mode provenance добавляется только в новый Stage9-клон accepted receipt, исходный Stage6 receipt не переписывается.

Для legacy без snapshot реализовано read-only восстановление исходных settings bytes в памяти из current exact SHA или history с exact original SHA, указанным в sealed manifest. Численного сходства недостаточно. Без evidence — LEGACY_AKZ_SETTINGS_UNPROVEN с понятным предложением создать новый TEST-комплект. Код/10 targeted source tests покрывают эту ветвь; **actual frozen legacy0.9.11 сценарий NOT_RUN после STOP**, успешность независимого восстановления на EXE пока не заявляется. Старые manifests/receipts не переписывались.

**6. OLD/NEW RUN VALIDATION.** Обязательный A–L выполнен в доступной CLI-области на собранной0.9.11 до STOP: A9/2026 с15–25; повторное сохранение всех параметров как float сохранило live SHA; повторная генерация всех четырёх вариантов PASS без изменения датированных климатических records. Реальные16–26 на тех же датах дали STOP. B10/2026 сформирован принятым Stage1 в той же свежей TEST-копии с новыми датами и snapshot16–26. A/B проверены после изменения live-настроек, в новых процессах EXE и даже при временно отсутствующем live settings.json. Native readonly Word/Excel verification обоих — PASS; каждый12 физических файлов/13 records (Word3/Excel9), М1/данные/SHA сохранены. Контролируемое повреждение только disposable TEST-snapshot A дало ожидаемый FROZEN_SETTINGS_SHA_CONFLICT; exact bytes/attributes восстановлены, повтор verify PASS. Manifests/receipts для fault injection не менялись.

A: `{ab['A_manifest']}`. B: `{ab['B_manifest']}`. Новые датыB: {', '.join(sorted(ab['new_B_dates'].values()))}. {link('AB_FINAL.json')}, {link('LIVE_SETTINGS_INDEPENDENCE.json')}.

Отдельный legacy fixture создан сохранённой0.9.10; во время его формирования зафиксирован foreign exit. Поэтому следующие0.9.11 шаги старого float-day reuse, exact-history recovery и missing-evidence STOP **NOT_RUN**. Сценарий нельзя считать завершённым.

**7. TARGETED TESTS.** Source **42/42** (32 АКЗ+10 persistence), Office regression **17/17 PASS**; SOURCE_42.log/OFFICE_17.log. Native source генерация всех четырёх вариантов PASS, отдельно от EXE. Frozen0.9.11: основной19-check сценарий и2 live-absent проверки PASS gates, включая3 ожидаемых FAIL/STOP. Каждый CLI стартует новый EXE-процесс с PATH только System32/Windows и без PYTHONPATH/PYTHONHOME. GUI clicks/relaunch не выполнены. Legacy6 новых EXE checks NOT_RUN. Текущий safety verdict STOP, общий полный validation не PASS.

Магнус, варианты0/1/2/3, слои1/1/2/3, материалы/цвета/нормы, ∑1+2=180 (без придуманной индивидуальной80), app formulas/format и корректность видимой ВОМ проверяются прежними tests/native validators. Официальный режим не включался. Stage8/backup source byte-identical; полный audit Stage1–8 и повтор незатронутых Stage8 operational tests NOT_RUN по targeted scope.

**8. EXE PATH / SHA.** `{EXE}`; SHA `{build['exe_sha256']}`. Product SHA (runtime source+spec map) `{product_sha}`. Onedir {build['file_count']} файлов,63 raw Python modules/52 compiled modules совпадают с исходниками. {link('BUILD_MANIFEST.json')}, {link('SOURCE_AT_BUILD.json')}, {link('BUILD_ENVIRONMENT.json')}. Сборка0.9.10 сохранена целиком: {len(old['files'])} файлов, SHA `{old['exe_sha256']}`. Новая EXE — **кандидат с незавершённой валидацией**, не принятый релиз.

**9. SOURCE INTEGRITY.** 56/56 Stage1–8 файлов неизменны;32 sealed DEV seed файла и8 реальных исходных образцов неизменны. ВОМ/реальные справочники/реестры/паспорта не записывались. В свежий TEST-стенд скопирован прежний artificial assignment fixture с тем же SHA без редактирования содержимого; новый B добавлен принятым Stage1 только в его TEST-реестр. A/B documents, manifests, receipts и snapshots имеют прежние SHA после всех fault injections. {link('SOURCE_INTEGRITY_FINAL.json')}.

**10. CHANGED FILES.**

{changed}

Также добавлены текущие validation-контроллеры/evidence. Неисполненный finish.py защищён от объявления завершения при FOREIGN_OFFICE_STOP; его шаблон не является отчётом о результате.

**11. NOT_RUN / REMAINING.** Legacy frozen0.9.11 recovery/reuse/missing-evidence checks; GUI Computer Use/E2E/viewer/relaunch; независимый targeted Work09; полный audit Stage1–8; production/official/deployment. Причина завершения чужого48680 UNKNOWN. Исторический Word3132 также UNKNOWN. Нужно разобраться с новым foreign exit и продолжить оставшуюся native-валидацию только после снятия STOP пользователем. Испытательный монитор требует fail-fast отмены собственного текущего задания: сейчас флаг потреблялся лишь после его окончания; это отдельное ограничение harness, не исправление ownership продукта.

**12. READY FOR TARGETED WORK 09: NO.** Два persistence исправления реализованы и A/B CLI проверены, но safety STOP и незавершённый legacy сценарий не позволяют объявить кандидат готовым. READY FOR USER E2E: NO. READY FOR PRODUCTION: NO. После отчёта STOP; независимый Work не запускался, deployment не выполнялся.
'''
(OUT/'REPORT.md').write_text(report,encoding='utf8')
(OUT/'WORK09_ASSIGNMENT.md').write_text(f'''DRAFT / DO NOT DISPATCH — targeted Work09 persistence re-review, candidate0.9.11.
READY FOR TARGETED WORK09: NO. Current stop: foreign baseline ExcelPID48680 exited, causeUNKNOWN. Read REPORT.md/FOREIGN_OFFICE_STOP.json/OFFICE_STOP_DIAGNOSIS.json first. No further native/GUI tests before user releases STOP. GUI additionally requires separate explicit permission.

Candidate `{EXE}`, SHA `{build['exe_sha256']}`; onedir1120, productSHA `{product_sha}`. Read EXACT_CODE_DIFF.patch, CHANGED_FILES.json, SOURCE_AT_BUILD.json, BUILD_MANIFEST.json, SOURCE_INTEGRITY_FINAL.json, AB_FINAL.json, LIVE_SETTINGS_INDEPENDENCE.json, source42/Office17 logs and command/process/ownership evidence. Stage1–8/production/real VOM/catalogs/templates/passports/old manifests/receipts must not change. No full audit, real deductions, official release or deployment.

Completed self-evidence: numeric canonicalization/comma/precision/no-opSave, frozen snapshot/manifest/receipt binding, A15–25→floatSave→repeat, same-date16–26 STOP, B newTEST10/2026/new dates, A/B native readonly verification, new EXE processes/live-file absent verify, frozen tamper STOP with exact TEST restoration. GUI NOT_RUN.

After explicit resumption, finish old0.9.10 float-day compatibility on new0.9.11, exact original settings SHA reconstruction from history, missing evidence UNPROVEN/newTEST suggestion, all old document/manifest/receipt/day SHA unchanged. Source tests cover the implementation but are not frozen legacy proof. Review readonly/no-overwrite/run-path binding, true parameter digest changes, no rounding, intact same-date rule/Magnus/norms/∑1+2=180, disabled-auto copied-receipt provenance. Also fix the TEST harness so foreign-exit signal stops dispatch and safely cancels only its proven own current job; do not introduce unknown-PID termination or change accepted ownership code for convenience.

Fresh Office baseline must use PID+creation time+image and retained read-only handles. Unexpected foreign exit →STOP/diagnostics. Current causeUNKNOWN must not be attributed without evidence; exit0 is not attribution. Old Word3132 remainsUNKNOWN. Report independent identity/source/EXE/product hashes, PASS/FAIL/NOT_RUN, findings and limitations; do not call self-tests independent. After report STOP.
''',encoding='utf8')
save(OUT/'RESULT.json',{'status':'STOP_FOREIGN_OFFICE_EXIT','version':'0.9.11','exe':str(EXE),'exe_sha256':build['exe_sha256'],'product_sha256':product_sha,'source_tests':42,'office_regression_source':17,'completed_0911_frozen_checks':21,'legacy_0911_checks':'NOT_RUN_STOP','GUI':'NOT_RUN','foreign_exit':stop,'READY_FOR_TARGETED_WORK_09':'NO','READY_FOR_USER_E2E':'NO','READY_FOR_PRODUCTION':'NO','report':str(OUT/'REPORT.md')})
print('STOP report saved; changed files',len(changes),'product SHA',product_sha,flush=True)
