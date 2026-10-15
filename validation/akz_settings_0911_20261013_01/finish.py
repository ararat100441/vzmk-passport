from common import *
if (OUT/'FOREIGN_OFFICE_STOP.json').exists():raise RuntimeError('STOP_FOREIGN_OFFICE_EXIT: completion report forbidden')
import difflib
ab=read(OUT/'AB_FINAL.json');legacy=read(OUT/'LEGACY_FINAL.json');build=read(OUT/'BUILD_MANIFEST.json')
assert all(sha(p)==h for p,h in read(OUT/'ACCEPTED_BEFORE.json').items())
assert all(sha(p)==h for p,h in read(OUT/'SOURCE_AT_BUILD.json').items())
for field in ('A_files','B_files'):assert all(sha(p)==h for p,h in ab[field].items())
assert all(sha(p)==h for p,h in legacy['old_documents_manifests_receipts'].items())
assert all(sha(p)==h for p,h in legacy['day_files_unchanged'].items())
old=read(ROOT/'validation/akz_auto_20261013_02/BUILD_MANIFEST.json');oldexe=Path(old['exe']);assert {str(p.relative_to(oldexe.parent)):sha(p) for p in oldexe.parent.rglob('*') if p.is_file()}==old['files']
assert {str(p.relative_to(EXE.parent)):sha(p) for p in EXE.parent.rglob('*') if p.is_file()}==build['files']
oldintegrity=read(ROOT/'validation/akz_auto_20261013_02/SOURCE_INTEGRITY_FINAL.json')
for ref in oldintegrity['real_first_passport_template_pairs']:assert sha(ref['path'])==ref['sha256']
seed=read(ROOT/'validation/akz_auto_20261013_02/SEED_BEFORE.json');assert all(sha(p)==h for p,h in seed['files'].items())
rows=[];diff=[]
for p in [ROOT/'stage9.py',*sorted((ROOT/'stage9_app').glob('*.py')),ROOT/'stage9_0911.spec',ROOT/'tests/stage9_akz_auto/test_persistence.py']:
    rel=p.relative_to(ROOT);before=OUT/'source_before'/rel;oldtext=before.read_text(encoding='utf8') if before.exists() else '';new=p.read_text(encoding='utf8')
    if oldtext!=new:
        rows.append({'path':str(p),'relative':str(rel),'before_sha256':sha(before) if before.exists() else None,'after_sha256':sha(p)})
        diff+=difflib.unified_diff(oldtext.splitlines(True),new.splitlines(True),fromfile='0.9.10/'+rel.as_posix(),tofile='0.9.11/'+rel.as_posix())
save(OUT/'CHANGED_FILES.json',rows);(OUT/'EXACT_CODE_DIFF.patch').write_text(''.join(diff),encoding='utf8')
product={str(Path(p).relative_to(ROOT)):h for p,h in read(OUT/'SOURCE_AT_BUILD.json').items() if not Path(p).is_relative_to(ROOT/'tests')}
product_sha=hashlib.sha256(json.dumps(product,ensure_ascii=False,sort_keys=True,separators=(',',':')).encode('utf8')).hexdigest()
save(OUT/'SOURCE_INTEGRITY_FINAL.json',{'status':'PASS','accepted_stage1_8_count':len(read(OUT/'ACCEPTED_BEFORE.json')),'accepted_stage1_8_unchanged':True,'seed_files_unchanged':len(seed['files']),'real_template_files_unchanged':8,'preserved_0910_exe_sha256':sha(oldexe),'preserved_0910_onedir_files':len(old['files']),'new_0911_onedir_files':len(build['files']),'product_sources':product,'product_sha256':product_sha,'old_manifests_receipts_unchanged':True,'A_B_unchanged':True,'real_VOM_catalogs_registry_passports_written':False,'TEST_fixture_note':'Previously accepted artificial assignment workbook copied byte-for-byte into fresh TEST workspace, no content edits; accepted Stage1 created new TEST passport10/2026 only in that workspace.'})
jobs=[];seen=set()
for p in sorted(OUT.glob('*.json')):
    d=read(p)
    if not isinstance(d,dict) or 'office_supervisor' not in d or d['office_supervisor'] in seen:continue
    seen.add(d['office_supervisor']);s=read(d['office_supervisor']);assert s['owned_orphans']==0 and not s['ownership_errors'];assert all(not i.get('forced_termination_attempted') and i['exit_code']==0 for i in s['instances'])
    command=read(p.with_suffix('.command.json'));version='0.9.11' if command['command'][0]==str(EXE) else '0.9.10' if command['command'][0]==SETUP['exe_0910'] else 'SOURCE'
    jobs.append({'result':str(p),'version':version,'status':d['status'],'supervisor':d['office_supervisor'],'office_instances':len(s['instances']),'orphans':0,'force':0})
    folder=Path(d['office_supervisor']).parent
    assert folder.is_relative_to(Path(SETUP['base'])/'VZMKPassport/diagnostics')
    for file in folder.rglob('*'):
        if file.is_file():
            target=OUT/'office_jobs'/folder.name/file.relative_to(folder);target.parent.mkdir(parents=True,exist_ok=True);shutil.copyfile(file,target);assert sha(file)==sha(target)
baseline=read(OUT/'OFFICE_BASELINE_FINAL.json');assert not baseline['foreign_stop'] and all(p['alive'] for p in baseline['processes'])
save(OUT/'OFFICE_FINAL.json',{'status':'PASS','jobs':jobs,'own_orphans':0,'forced_terminations':0,'foreign_baseline_preserved':baseline,'0911_tasks':sum(j['version']=='0.9.11' for j in jobs),'0911_office_instances':sum(j['office_instances'] for j in jobs if j['version']=='0.9.11')})
def link(name):return f'[{name}]({(OUT/name).as_posix()})'
change_text='\n'.join(f"- `{r['relative']}` — SHA `{r['after_sha256']}`" for r in rows)
report=f'''ВЗМК — Stage 9, targeted settings persistence fix. DEV EXE 0.9.11.

**1. ROOT CAUSE MAJOR 1.** В 0.9.10 GUI вызывал float(), а daily_climate использовал S.digest сырого JSON. 15 и 15.0 давали разные bytes/digest при одинаковых диапазонах. Подтверждение независимого Work: `validation/work09_akz_auto_0910_20261013_01/SETTINGS_REPRESENTATION_AND_MANIFEST_DEFECT.json`. Исправление локализовано в Stage9; общий S.digest и принятые алгоритмы не менялись.

**2. ROOT CAUSE MAJOR 2.** _generate_existing переносил изменяемый preferences/akz_auto/settings.json из preview fingerprints в manifest.inputs. verify_manifest сравнивал старый SHA с текущим файлом. Поэтому легитимная смена настроек инвалидировала неизменённые документы. Подтверждение Work: MANIFEST_MUTABLE_SETTINGS_CONFIRMED.json.

**3. EXACT CODE DIFF.** {link('EXACT_CODE_DIFF.patch')}; полный список before/after SHA: {link('CHANGED_FILES.json')}. Новый helper akz_settings.py обслуживает только настройки Stage9. Выбор документов, формула Магнуса, нормы, ВОМ, Stage1–8, ownership и Office lifecycle не менялись.

**4. NUMERIC CANONICALIZATION.** Детерминированный JSON: отсортированные ключи, фиксированные разделители, конечные числовые значения без экспоненциальной записи/лишних нулей, −0 эквивалентен 0. 15/15.0/15.00 равны; 15.1 и 15.100000000000001 различны, округления нет. GUI принимает десятичную запятую через Decimal и отвергает преобразование с потерей точности. Прежние границы и точность диапазонов сохранены. Численно равное сохранение — no-op, сохраняет исходные bytes/SHA и history. Настоящая смена диапазона даёт новый digest.

Для старых климатических записей без версии digest проверяются только точные int/float представления тех же шести границ, максимум64 комбинации. Seal, значения, формула, диапазоны и Td-margin проверяются. Старые записи не переписываются и не генерируются заново. Новые записи содержат canonical digest/schema и климатические настройки. Реальная смена 15–25→16–26 на прежней дате по-прежнему STOP.

**5. FROZEN SETTINGS.** До формирования нового общего комплекта создаётся run/inputs/AKZ_SETTINGS.json с точными bytes рабочего файла, SHA-256 и canonical digest; файл read-only, повторное создание/перезапись запрещены. Manifest.inputs содержит snapshot и SHA receipts, исключая live settings.json. Manifest и PAIR receipts связаны с одним reference. Для выключенного auto-режима provenance добавляется только в новую Stage9-копию accepted receipt; оригинал Stage6 не меняется. Standalone auto-generation также создаёт run snapshot и sealed COMPLETE с hashes receipts.

Проверка требует snapshot именно этого owned run, правильные path/SHA/digest/schema, schema/нормы настроек, связь всех receipts и SHA документов. Подмена другим run или повреждение — STOP. Новые комплекты независимы от live settings.json. Старые manifests/receipts не переписываются. Для legacy используется текущий точный SHA либо восстановление исходных bytes из history в памяти по SHA, записанному в исходном sealed manifest. Численное сходство без точного SHA недостаточно. Без evidence — LEGACY_AKZ_SETTINGS_UNPROVEN и понятное предложение создать новый TEST-комплект.

**6. OLD/NEW RUN VALIDATION.** Обязательный A–L сценарий выполнен на собранной EXE 0.9.11, GUI не использовался. A: паспорт9/2026, воздух15–25. Сохранение всех GUI-диапазонов как float не изменило live SHA; повторная генерация всех четырёх вариантов PASS, исходный климат по датам неизменен. Настройки16–26 на старых датах дали STOP. B: новый искусственный паспорт10/2026, созданный принятым Stage1 в свежем TEST-стенде, отдельный run и новые даты; snapshot16–26. A/B повторно проверены после изменения live-файла, в новых процессах EXE и даже при временно отсутствующем live-файле. A/B сохраняют собственные настройки, docs/SHA, receipts и manifest. Native Word/Excel read-only verification обоих полных комплектов PASS (каждый12 физических файлов/13 records, Word3/Excel9). {link('AB_FINAL.json')}, {link('LIVE_SETTINGS_INDEPENDENCE.json')}.

A: `{ab['A_manifest']}`. B: `{ab['B_manifest']}`. Новые даты B: {', '.join(sorted(ab['new_B_dates'].values()))}. Контролируемое повреждение только TEST-snapshot A дало AKZ_FROZEN_SETTINGS_SHA_CONFLICT; точные bytes/attributes восстановлены, повторная проверка PASS. Ни один manifest/receipt для fault injection не менялся.

Legacy-комплект создан сохранённой EXE0.9.10 в отдельной свежей TEST-копии с float-климатом. EXE0.9.11 повторно использовала старые day-records без изменения SHA. После реальной смены live-настроек старый manifest проверен по точному исходному SHA из history. При временно недоступном единственном исходном evidence получен ожидаемый понятный STOP; evidence возвращён с прежними bytes/path, restart verify PASS. Старые документы, manifests, receipts и day-records неизменны. {link('LEGACY_FINAL.json')}.

**7. TARGETED TESTS.** Source42/42 (32 прежних АКЗ +10 persistence), Office regression17/17 PASS. Логи SOURCE_42.log/OFFICE_17.log. Source native generation всех вариантов PASS — отдельно от EXE. EXE: основной сценарий19 проверок, live-file independence2, legacy6; ожидаемые FAIL/STOP записаны как FAIL, не скрыты. Итоговые gates PASS. CLI каждого шага запускается отдельным процессом с PATH только System32/Windows, PYTHONPATH/PYTHONHOME отсутствуют. Сохранены команды, runtime/executable SHA, process results и Office supervisor evidence.

Регрессия Магнуса, всех вариантов, материалов/цветов/слоёв/норм, ∑1+2=180, формул приложений, структуры/оформления и ВОМ проходит через прежние source tests и native generator validators. Новые A/B дополнительно открыты native Word/Excel read-only, CorruptLoad=0; М1, numbers/dates и SHA проверены. Stage8/backup source byte-identical; полный Stage1–8 audit и повторные незатронутые Stage8 operational tests **NOT_RUN** по targeted scope. Принятый Stage1 использован только для допустимого нового TEST-паспорта B. Реальные списания/закрытие/official mode/deployment не выполнялись.

**8. EXE PATH / SHA.** `{EXE}`. SHA-256 `{build['exe_sha256']}`. Product SHA (детерминированный runtime source+spec map) `{product_sha}`. Onedir {build['file_count']} файлов; 63 raw Python modules и52 compiled modules соответствуют исходникам. {link('BUILD_MANIFEST.json')}, {link('SOURCE_AT_BUILD.json')}, {link('BUILD_ENVIRONMENT.json')}. Полная0.9.10 сохранена: {len(old['files'])} файлов, SHA `{old['exe_sha256']}`. Внешний Python не требуется; native Word/Excel установлен на тестовом Windows.

**9. SOURCE INTEGRITY.** 56/56 принятых Stage1–8 файлов неизменны; sealed DEV seed32 файла и8 реальных исходных образцов неизменны. Реальные ВОМ/справочники/реестры/готовые паспорта не записывались. В свежие TEST-копии только скопирован прежний искусственный DEV-справочник назначений без редактирования его содержимого; B добавлен принятым Stage1 только в собственный TEST-реестр. {link('SOURCE_INTEGRITY_FINAL.json')}.

**10. CHANGED FILES.**

{change_text}

Также добавлены текущие validation-контроллеры/evidence. Старые отчёты, manifests, receipts и сборки не исправлялись задним числом.

**11. NOT_RUN.** GUI Computer Use/GUI E2E/GUI viewer clicks и GUI-relaunch0.9.11; независимый targeted Work09 review; production/official issuance/deployment; полный audit Stage1–8. GUI Save покрыт headless проверками GUI parser и того же backend, а не выдаётся за реальный click test. Новый baseline Office сохранён: {', '.join(str(p['ProcessId']) for p in baseline['processes'])}; retained read-only handles живы с прежней creation time, собственных orphan/force0. {link('OFFICE_FINAL.json')}. Старый WordPID3132 остаётся историческим UNKNOWN.

**12. READY FOR TARGETED WORK 09: YES.** Это готовность к независимому review кандидатной сборки после собственных проверок, а не положительный независимый verdict. READY FOR USER E2E: NO. READY FOR PRODUCTION: NO. {link('WORK09_ASSIGNMENT.md')}. После отчёта STOP.
'''
(OUT/'REPORT.md').write_text(report,encoding='utf8')
(OUT/'WORK09_ASSIGNMENT.md').write_text(f'''Targeted Work09: re-review ONLY two persistence MAJOR findings, candidate0.9.11.

EXE `{EXE}`, SHA `{build['exe_sha256']}`; full onedir {build['file_count']} files. Read REPORT.md, EXACT_CODE_DIFF.patch, CHANGED_FILES.json, SOURCE_AT_BUILD.json, BUILD_MANIFEST.json, SOURCE_INTEGRITY_FINAL.json, AB_FINAL.json, LEGACY_FINAL.json, LIVE_SETTINGS_INDEPENDENCE.json, source42/Office17 logs, commands/process/supervisor evidence. Verify evidence manifest SHA before review. No full Stage1–8 audit; no production/deployment/official release or real registry/VOM/catalog/template/passport edits. All tests on new isolated copies. GUI Computer Use requires separate user permission; otherwise NOT_RUN. Do not run Work or message others from this assignment without user authority.

MAJOR1: independently inspect deterministic numeric JSON and GUI comma parser. 15/15.0/15.00 must give the same digest/live SHA on equivalent save, no rounding of genuine decimal differences. Test valid/invalid precision, real16–26 change, old integer/float/mixed day digest compatibility, unchanged old day-record SHA, same-date real conflict STOP. Never relax shared Stage8 digest or silently regenerate climate.

MAJOR2: inspect owned run snapshot copied at generation start, exact SHA/canonical digest, readonly/no-overwrite, live setting race/stale preview STOP, path bound to run, manifest.inputs excluding live setting and pinning snapshots/receipts. Receipts must bind the same frozen snapshot. Disabled-auto provenance must affect only new Stage9 copies of accepted receipts. Verify old/new runs with future/live-absent settings; corrupt/substitute snapshots/receipts only in disposable TEST fixtures and require STOP. Old manifests/receipts must remain byte-identical. Legacy recovery is exact original SHA from current bytes or reconstructed history bytes in memory; missing evidence must show clear UNPROVEN +new TEST suggestion. No unsupported legacy PASS.

Repeat mandatory A–L using CLI commands from evidence: A15–25, equivalent float Save, repeat; real16–26 on A date STOP; B newTEST10/2026 with accepted Stage1/new dates; re-open A/B and compare frozen settings/docs/SHA; fresh EXE processes verify both. Actual GUI clicks are NOT_RUN unless separately authorized. Native Excel/Word readonly normal opens on copies only. Retain Magnus, variants0/1/2/3, layers1/1/2/3, materials/norms, cumulative∑1+2=180 (no invented individual80), visible VOM/М1, app formulas/formatting and official STOP. Compare Stage1–8 hashes; do not change accepted algorithms.

Before any native tests capture fresh Office identities (PID+creation time+image) and preserve foreign processes with read-only handles. Unexpected foreign exit →immediate STOP/diagnostics. Only proven own workers may be cleaned; no global taskkill/UAC/unknownPID termination. Historical Word3132 cause UNKNOWN. Do not present Codex source tests as frozen CLI or GUI, or self-tests as independent Work evidence.

Return independent verdict, findings with exact source/field/proof, source/EXE/product SHA, old/new run validation, numeric and frozen/legacy tests, native Office/integrity, FAIL/NOT_RUN and remaining limits. After report STOP. No deployment.
''',encoding='utf8')
save(OUT/'RESULT.json',{'status':'PASS_TARGETED_SELF_VALIDATION','version':'0.9.11','exe':str(EXE),'exe_sha256':build['exe_sha256'],'product_sha256':product_sha,'source_tests':42,'office_regression_source_tests':17,'mandatory_frozen_checks':19,'live_absent_checks':2,'legacy_new_exe_checks':6,'GUI':'NOT_RUN','independent_work':'NOT_RUN','READY_FOR_TARGETED_WORK_09':'YES','READY_FOR_USER_E2E':'NO','READY_FOR_PRODUCTION':'NO','report':str(OUT/'REPORT.md')})
print('Final report/integrity prepared; changed files',len(rows),'product SHA',product_sha,flush=True)
