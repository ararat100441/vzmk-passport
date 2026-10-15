from review import *
from openpyxl import load_workbook
m=json.loads((IMPL/'REVIEW_MANIFEST.json').read_text(encoding='utf-8'))
before=json.loads((OUT/'protected_before.json').read_text(encoding='utf-8'))
after={p:sha(p) if Path(p).is_file() else 'MISSING' for p in before}
changes={p:{'before':before[p],'after':h} for p,h in after.items() if h!=before[p]}
assert not changes,changes
write('protected_after.json',after)
native=json.loads((OUT/'native_exact_six.json').read_text(encoding='utf-8'))
assert len(native)==6 and all(x['unchanged'] and x['ReadOnly'] and x['AutomationSecurity']==3 and not x['EnableEvents'] for x in native)
preview=json.loads((OUT/'independent_preview.json').read_text(encoding='utf-8'))
contract=json.loads((ROOT/'validation/stage6_new_templates_audit_20261009_125930/IMPLEMENTATION_CONTRACT.json').read_text(encoding='utf-8'))
for x in native:
    n=x['variant'];s=x['sheets'][0]['cells_after_full_rebuild'];g=preview['groups'][str(n)]['payload']
    assert x['sha_before']==x['sha_after'] and x['sheets'][0]['pages']==1
    if x['kind']=='act':
        for cell,key in [('E5','contract'),('E6','customer'),('E7','object'),('B44','cipher')]:assert s[cell]==g['passport'][key]
        assert s['G8']==f'9/2026-{n} от ' and s['I8']==46300
        for layer in contract['variants'][str(n)]['layers']:
            for field in ('material','color'):assert s[layer[field]['cell']]==layer[field]['value']
            assert s[layer['required_thickness']['cell']]==layer['required_thickness']['raw']
            for cell in layer['historical_fields'].values():assert s[cell]=='НЕТ ДАННЫХ / ТЕСТ'
    else:
        r=g['selected'][0]
        for col,field in zip('ABCDEFG',('mark','name','quantity','mass_unit','mass_total','area_unit','area_total')):
            assert abs(float(s[col+'5'])-float(r[field]))<1e-8 if col not in 'AB' else s[col+'5']==r[field]
        assert s['A2']=='Объект: '+g['passport']['object']
    assert not any(isinstance(v,str) and v.startswith(('#REF!','#VALUE!','#DIV/0!','#NAME?')) for sheet in x['sheets'] for v in sheet['cells_after_full_rebuild'].values())
safety=json.loads((OUT/'independent_safety.json').read_text(encoding='utf-8'));assert len(safety)==13 and all(x['status']=='PASS' for x in safety)
assert 'Ran 33 tests' in (OUT/'independent_unit_tests.txt').read_text(encoding='utf-8')
assert 'OK' in (OUT/'independent_unit_tests.txt').read_text(encoding='utf-8')
original=json.loads((IMPL/'targeted_validation.json').read_text(encoding='utf-8'));assert len(original['results'])==14 and all(x['status']=='PASS' for x in original['results'])
for p,h in m['product_current_sha256'].items():assert sha(p)==h
for x in m['six_outputs']:assert sha(x['path'])==x['sha256']
for p,h in m['receipts_sha256'].items():assert sha(p)==h
for key in ('tested_before_sha256','tested_after_sha256'):assert all(sha(p)==h for p,h in m[key].items())
for p,r in m['evidence_files'].items():assert sha(p)==r['sha256'] and Path(p).stat().st_mtime_ns<=(IMPL/'REVIEW_MANIFEST.json').stat().st_mtime_ns
for x in native:
    suffix='act_entire_workbook' if x['variant']==2 and x['kind']=='act' else x['kind']
    assert (OUT/f'visual_{x["variant"]}_{suffix}.png').is_file()
write('integrity.json',{'protected_count':len(before),'changes':changes,'product_match':9,'exact_output_match':6,'native_full_rebuild_value_assertions':'PASS','implementation_targeted_results':14,'independent_safety_results':13,'independent_unit_results':33,'actual_Tk_preview_and_R2':'PASS','production_read_write':'NONE by bounded code and Python open guard; no OS-wide trace','network_calls':'NONE requested; Python socket guard; no OS-wide trace','manifest_last_implementation':'PASS'})
lines=[]
def section(title,text):lines.extend([f'## {title}','',text,''])
section('1. VERDICT — PASS','Независимая проверка Stage 6 Part 2 завершена. Проверен текущий продукт и exact шесть TEST XLSX. BLOCKER 0 / MAJOR 0 / MINOR 0. Официальная выдача остаётся STOP.')
section('2. ACTUAL STATE','Product set: 9 файлов; TESTED_BEFORE = TESTED_AFTER = CURRENT для шести tested файлов, оставшиеся три accepted файла неизменны. Фактически изменены gui.py и STAGE6_DEV.md; добавлены part2.py и test_stage6_part2.py. Других product changes относительно source snapshot не обнаружено.\n\n| Product exact path | SHA256 |\n|---|---|\n'+'\n'.join(f'| {p} | `{h}` |' for p,h in m['product_current_sha256'].items())+'\n\n| Exact TEST output path | SHA256 |\n|---|---|\n'+'\n'.join(f'| {x["path"]} | `{x["sha256"]}` |' for x in m['six_outputs'])+'\n\nSource fingerprints (текущие и неизменные):\n\n| Source exact path | SHA256 |\n|---|---|\n'+'\n'.join(f'| {p} | `{h}` |' for p,h in m['source_after_sha256'].items() if Path(p).suffix.lower() in ('.xls','.xlsx','.xlsm') and (Path(p).name in ('ВЗМК_DEV.xlsx','УЗК_DEV.xlsm','Профили.xlsx','Справочник марок.xlsx','Акт № 4-2026.xlsx','1_2026 - Приложение к АКЗ.xlsx','АКЗ-1.xlsx','АКЗ-2.xlsx','АКЗ-3.xlsx','Приложение к АКЗ-1.xlsx','Приложение к АКЗ-2.xlsx','Приложение к АКЗ-3.xlsx') or Path(p).name=='ВОМ отгрузочный счет 183  (5).xls'))+f'\n\nDEV assignment fixture: {m["fixture"]["path"]}; SHA `{m["fixture"]["sha256"]}`. DEV registry/UZK hashes соответствуют собственным checkpoints. Полный guard: protected_before.json / protected_after.json ({len(before)} путей), исходный relevant source set — 534 пути.')
section('3. TEMPLATE / GROUP MAPPING','Explicit Layout №1/2/3 → АКЗ-{n}.xlsx + Приложение к АКЗ-{n}.xlsx в exact папке Гараж казахстан/Паспорт. Все шесть pins совпали с независимым source contract. Fuzzy fallback отсутствует; missing/changed exact path останавливается. Из отдельного DEV fixture получены Бп2-5→1, Бп2-18→2, Бп2-19→3. Blank отличается от 0; duplicate assignment STOP. Номера 9/2026-1/-2/-3; дата 05.10.2026. Проверка isolated registry row401/O401=04.10.2026 подтвердила динамический lookup, без row400 hardcode.')
section('4. АКЗ-1 ASSESSMENT','PASS. Один слой E: Norecoat FD Primer Grey; 150 мкм; серый. RAL не придуман. Текущие договор 06/26 от 03.06.2026, заказчик ООО "СТК ВК", объект парковки и КМД WAA0007144-AC2-КМД подтверждены stored/native. Явная TEST-маркировка. Чужие реквизиты отсутствуют.')
section('5. АКЗ-2 ASSESSMENT','PASS. Два слоя E/I, ordinals E17/H17: Триокор Мастик 4500 / Триокор Финиш 5500; Серый / 7015; E25=100; I25 буквально ∑1+2=180. Preview required_individual_microns=None / NOT_PROVEN для второго слоя; самостоятельные 80 не выведены. Старые 115/220 отсутствуют.')
section('6. АКЗ-3 ASSESSMENT','PASS. E/H/J: Триокор Мастик 4500 / Триокор Мастик 4500 / Триокор Финиш 5500; 100/80/60 мкм; Серый/Серый/7015. Старые 115/220/310 отсутствуют. Номер, дата, текущие реквизиты и TEST-маркировка подтверждены.')
section('7. APPLICATIONS / TOTALS','PASS. Selected current shipment VOM read-only: 104 visible data rows, 397 hidden excluded. A:G: mark/name/quantity/unit mass/total mass/unit area/total area; исходный порядок и повторные строки сохраняются.\n\n| АКЗ | VOM row | Марка | Кол-во | Единичная = общая масса, кг | Единичная = общая площадь, м² | Итого |\n|---|---|---|---|---|---|---|\n| 1 | 40 | Бп2-5 | 1 | 37.20 | 1.412 | 66 |\n| 2 | 53 | Бп2-18 | 1 | 31.32 | 1.184 | 52 |\n| 3 | 54 | Бп2-19 | 1 | 36.74 | 1.393 | 99 |\n\nUnused body clear+hidden; total row visible; три исходные SUM и объединения сохранены. Native full rebuild дал правильные итоги. Площадь взята из выбранной ВОМ, старые приложения числовым источником не служат.')
section('8. CAPACITY / SAFETY','PASS. Independent isolated native generation: 61 PASS/62 STOP; 47 PASS/48 STOP; 94 PASS/95 STOP. Все строки на границе сохранены; превышение остановлено до copy/publication. Проверены изменение каждого из пяти inputs после preview, existing-output STOP с сохранением USER.txt, actual-save первой книги + injected failure второй + полный rollback. Native Windows lock второй книги дал FILE_LOCKED/WinError32; исходный объект ошибки сохранён, cleanup secondary typed; живой handle держался до завершения вызова. При lock/foreign directory/foreign file/replaced inode cleanup не удалял чужие данные. После закрытия lock fresh-process recovery и повторная генерация прошли. Failed group → PARTIAL_STOP; preflight invalid group блокирует все выбранные до публикации. 13 independent safety cases PASS; 33 unit cases rerun PASS. Mutation fixtures/output только в новом review каталоге.')
section('9. HISTORICAL MEASUREMENT SAFETY','PASS. Проверены все output sheets, весь xl XML, исходные неизменённые nonempty cells и native значения после CalculateFullRebuild. Климат, партии, даты/время нанесения, wet/dry/total measured thickness, приборы, комиссия и подписи очищены либо явно НЕТ ДАННЫХ / ТЕСТ; selector marks — ?. Безусловное историческое подтверждение соответствия заменено TEST-запретом выдачи. Формулы актов только G8=E4/I8=H4; приложений только три SUM. Нет RAND/RANDBETWEEN, live dew-point calculation, #REF!, внешних связей, embedded/VBA/media изображений подписей. Official mode останавливается до источников/выхода.')
section('10. CLIMATE / EXTRA SHEET','PASS. Один общий ClimateDay на 2026-10-05 для всех трёх групп, все readings=None и confirmation=MISSING. Нет случайных показаний по слоям/актам и фиктивного surface confirmation. Td-алгоритм не выбран. АКЗ-2 Лист1 veryHidden; единственное значение A1 — явный служебный TEST/MISSING. Никаких зависимостей от исторических helper данных. Native предпросмотр Напечатать всю книгу подтвердил 1 из 1, без служебной страницы.')
section('11. NATIVE EXCEL VISUAL VALIDATION','PASS. Exact шесть SHA opened independently в Microsoft Excel 16.0: ReadOnly=True, UpdateLinks=0, AutomationSecurity=3, EnableEvents=False, CorruptLoad=0 (normal load), CalculateFullRebuild, Close(SaveChanges=False). Без repair/recovery диалогов; Save/экспорт/PDF/PrintOut не вызывались. Для всех шести просмотрены нормальный вид и Print backstage; сохранены visual_*.png. Текущие реквизиты, слои/материалы/требования, отсутствие псевдоизмерений и итоги проверены по native Value2. Каждый exact текущий документ имеет 1 страницу; нет обрезанного отдельного листа/лишней служебной страницы. Final native SHA AFTER = BEFORE для всех шести.')
section('12. №0 / R1 / R2 REGRESSION','PASS. Accepted core.py SHA eede991cf2d2ca259f251ef9f79beb49542ea1e536ff2eea6f18cf9c23cd935d неизменен. Exact accepted NEW_TEST_9_2026 две книги и PAIR.json unchanged, fingerprints в additional_checks.json. №0 mapping принятой пары сохранён. Новый isolated N3: строки7:20 hidden, итог21 visible, totals3/105.26/3.989; номер/DEV-дата/selected VOM areas сохранены. Native R2 lock/secondary cleanup/fresh-process retry повторены. Actual Tk generate button handler текущего launcher сохранил первичную typed ошибку, инвалидировал preview и не сообщил успех. Полный повторный review №0 не проводился; реальные старые пары не пересоздавались.')
section('13. SOURCE / EVIDENCE INTEGRITY',f'PASS. {len(before)} protected paths BEFORE=AFTER; original 534 source fingerprints current match. Stage1–5, DEV registry/UZK, Stage4 profiles, реальный catalog, selected VOM, шаблоны0/1/2/3 и previous accepted evidence unchanged. Actual product delta: только четыре ожидаемых файла. Implementation 14 targeted PASS / 33 unit PASS journals и тестовый код прочитаны; current hashes связаны с tested before/after, receipts и exact outputs. Independent preview seal совпал с receipt act_id каждой пары. Все implementation evidence SHA и manifest-last проверены. Исторические incomplete attempts не используются для PASS. Production файлы не открывались/не хешировались; network/socket calls не выполнялись. Python guard запрещал production filenames и socket operations; OS-wide tracing не выполнялся. Новые записи только в isolated review папке. Собственный REVIEW_MANIFEST.json записывается после всех evidence/report и проверяется read-only.')
section('14. FINDINGS','BLOCKER 0\n\nMAJOR 0\n\nMINOR 0\n\nКонкретных нарушений проверяемого контракта не обнаружено.')
section('15. NOT_RUN','User E2E, официальная выдача, production, Stage7, физическая печать/PDF/export/deploy, OS-wide network/COM trace, полный Stage1–5 regression и полный повторный review №0 — NOT_RUN. Эти действия не требовались для данного independent TEST review; native Excel, визуальная проверка exact шести и обязательные isolated boundaries выполнены.')
section('16. REQUIRED FIX','NONE.')
section('17. READY FOR USER E2E — YES','Готовы exact три TEST-пары. После review STOP. Product/source/DEV outputs не исправлялись и не пересоздавались; Stage7 не начинался.')
(OUT/'FINAL_REPORT.md').write_text('\n'.join(lines),encoding='utf-8')
files={str(p):{'sha256':sha(p),'bytes':p.stat().st_size,'mtime_ns':p.stat().st_mtime_ns} for p in OUT.rglob('*') if p.is_file() and p.name!='REVIEW_MANIFEST.json'}
manifest={'schema':'STAGE6_PART2_INDEPENDENT_REVIEW_V1','created_utc':datetime.datetime.now(datetime.timezone.utc).isoformat(),'verdict':'PASS','findings':{'BLOCKER':0,'MAJOR':0,'MINOR':0},'ready_for_user_e2e':'YES','official':'STOP','product':m['product_current_sha256'],'exact_six':m['six_outputs'],'protected_sha256':after,'files':files,'manifest_last':'last write; self hash excluded'}
write('REVIEW_MANIFEST.json',manifest)
last=(OUT/'REVIEW_MANIFEST.json').stat().st_mtime_ns
assert all(sha(p)==r['sha256'] and Path(p).stat().st_mtime_ns<=last for p,r in files.items())
assert all(sha(p)==h for p,h in after.items())
print('FINAL_REVIEW_PASS',len(files),'evidence files; MANIFEST_LAST PASS; PROTECTED',len(after),'UNCHANGED',flush=True)
