import json,hashlib,os
from pathlib import Path
OUT=Path(__file__).resolve().parent;ROOT=OUT.parents[1]
def sha(p):return hashlib.sha256(p.read_bytes()).hexdigest()
def load(p):return json.loads(p.read_text(encoding='utf-8-sig'))
def save(n,d):(OUT/n).write_text(json.dumps(d,ensure_ascii=False,indent=2,default=str),encoding='utf8')
preview_path=ROOT/'validation/stage5_sample_results_20261008_130441/current_preview.json'
receipt_path=ROOT/'validation/stage5_sample_results_20261008_130441/new_document.json'
p=load(preview_path);receipt=load(receipt_path)
checks=[]
for name,expected in p['preview']['fingerprints'].items():
    path=Path(name);actual=sha(path) if path.is_file() else None
    checks.append({'path':name,'expected_sha256':expected,'actual_sha256':actual,'match':expected==actual,'receipt_field':'current_preview.json:preview.fingerprints'})
drift=[c for c in checks if not c['match']]
product=[]
for name,expected in receipt['product_binding'].items():
    actual=sha(Path(name)) if Path(name).is_file() else None
    if expected!=actual:product.append({'path':name,'expected_sha256':expected,'actual_sha256':actual,'receipt_field':'new_document.json:product_binding'})
assert len(drift)==3 and any(Path(c['path']).name=='inputs.py' for c in product)
output=Path(receipt['output']);output_sha=sha(output)
assert output_sha==receipt['sha256']=='8d3a2769baf6d381a0675804ef7f0806896f1b611e1498ea58677d3fdac1fd1b'
save('UZK_INPUT_BINDING_STOP.json',{'state':'STOP_UZK_ACCEPTED_INPUT_BINDING_MISMATCH','receipt_path':str(receipt_path),'receipt_sha256':sha(receipt_path),'preview_path':str(preview_path),'preview_sha256':sha(preview_path),'output_path':str(output),'output_sha256':output_sha,'output_sha_matches':True,'mode':receipt['mode'],'results_confirmed':receipt['results_confirmed'],'preview_seal_matches_receipt':p['seal']==receipt['preview_seal'],'input_checks':checks,'mismatches':drift,'generation_receipt_product_mismatches':product,'action':'STOP before Stage7 product edits, new VIK, quality generation, source reclassification or receipt rewrite'})
before=load(OUT/'SOURCE_BEFORE.json');after={n:sha(Path(n)) if Path(n).is_file() else None for n in before}
changes=[{'path':n,'before':h,'after':after[n]} for n,h in before.items() if after[n]!=h]
save('SOURCE_AFTER.json',after)
save('SOURCE_INTEGRITY.json',{'before_count':len(before),'after_existing_count':sum(h is not None for h in after.values()),'changes':changes,'pass':not changes,'scope':'Pasport product, DEV, existing E2E and project passport inputs; validation, cache/reparse, PDFs excluded; production registry never accessed','stage1_6_code_mutations':[]})
preflight=load(OUT/'INPUT_PREFLIGHT.json');apps=load(OUT/'APPENDIX_CANDIDATES.json');vom=load(OUT/'VOM_CURRENT.json')
table='\n'.join(f'| `{d["path"]}` | `{d["expected_sha256"]}` | `{d["actual_sha256"]}` |' for d in drift)
report=f'''# Этап 7 Implementation preflight STOP

09.10.2026, Europe/Samara. Реализация остановлена на обязательной проверке входов, до изменения product code и создания документов.

## 1. ACTUAL STATE

Root: `{ROOT}`. Git executable установлен, но `git rev-parse --show-toplevel` подтвердил отсутствие Git-репозитория в Pasport и родительских каталогах. Индексные инструменты недоступны; принятый код просмотрен точечно. Обязательные Stage7 audit report, field map, manifest design и OOXML evidence прочитаны. Только DEV, production registry не открывался.

DEV SHA `{preflight['registry_sha256']}` совпадает с ожидаемым. `read_registry()`/`select_project()` и динамический поиск проекта+номера подтвердили единственную строку400 для9/2026. Template SHA `{preflight['template_sha256']}` совпадает.

## 2. USER DECISIONS APPLIED

Зафиксированы короткий пункт14 и разрешение собрать конкретный TEST-комплект: существующее Приложение№1, новый ВИК из текущего DEV, выбранный SAMPLE_ONLY УЗК, исправленная пара АКЗ0. АКЗ1/2/3 исключены. Разрешение не отменяет пользовательские условия STOP и не разрешает переписывать старые receipts.

## 3. VERIFIED INPUTS

ВOM `{vom['path']}`:104 видимые строки,397 скрытых,288 изделий,15822,14кг. SHA `{preflight['vom_sha256']}`. Native Excel открыт read-only, макросы/ссылки отключены, сохранения нет.

Приложение-кандидат `{apps[0]['path']}`, SHA `{apps[0]['sha256']}`: номер9/2026, дата06.10.2026, объект совпадает. В нем123 строковых записи, в выбранной ВОМ104; часть марок разделена на несколько записей. Полная сверка по каждой марке (количество+масса), а не только общего итога, дала совпадение; differences={{}}. Никакой перегруппировки/изменения файла не выполнялось. Старое Stage1 E2E приложение с подтвержденным journal SHA `{apps[1]['sha256']}` относится к другому составу:50шт.,15071,44кг, и отвергнуто для текущей отгрузки. Новое приложение не создавалось. У текущего проектного приложения не приписан чужой original generator run_id.

Выбранная исправленная пара АКЗ0 найдена по указанному fix evidence и new_pair.json, оба файла/PAIR.json существуют, ожидаемые SHA обоих XLSX совпадают. Статус COMPLETE_TEST_PAIR, номер9/2026, дата05.10.2026, act_id=e9fc12e4433feafd. В существующем preview привязанные источники совпадают; TEST назначение0 только Бп2-5,Бп2-18,Бп2-19,3 изделия,105,26кг. Новая native перепроверка внутреннего содержимого АКЗ не выполнена после наступления STOP.

## 4. COHERENT TEST ASSEMBLY

**NOT_ASSEMBLED / STOP.** Новый assembly_run_id и frozen selected manifest не создавались. Нельзя объявить набор согласованным при неподтвержденной текущей input binding УЗК. Существующие receipts и source run IDs не изменены.

Кандидат УЗК: source cycle `stage5_sample_results_20261008_130441`, receipt `{receipt_path}`. Кандидат АКЗ: fix cycle `stage6_fix_r1_r2_20261009_094707`, ACT-ID e9fc12e4433feafd, receipt `{preflight['akz_folder']}\\PAIR.json`. Новый ВИК отсутствует. Никакие случайные/старые/FAILED/SUPERSEDED outputs не включены в новый manifest.

## 5. CURRENT DEV VIK

**NOT_RUN.** Старый Stage2 E2E не использован. Принятый `generate_vik()` просмотрен: API требует изолированный validation root, копию registry/documents, существующую папку текущего паспорта и запрещает overwrite. Безопасная генерация не проверялась на практике, поскольку предварительный УЗК gate завершился STOP. Stage2 код не изменялся.

## 6. UZK / AKZ0 SOURCE VALIDATION

УЗК DOCX существует: `{output}`. SHA `{output_sha}` совпадает с заданием и generation receipt. Receipt mode=SAMPLE_ONLY, results_confirmed=false; preview seal совпадает с receipt. Это **не повреждение самого DOCX**.

Однако текущие входы не совпадают с сохраненной привязкой принятого цикла:

| Вход | Expected из принятого preview | Actual сейчас |
| --- | --- | --- |
{table}

Для `uzk_stage5/inputs.py` расхождение также подтверждено непосредственно `new_document.json:product_binding`, а не только preview. Справочник и `passport_app/catalog.py` привязаны в `current_preview.json:preview.fingerprints`. Проверено {len(checks)} входных fingerprint-записей, расхождений3. В product_binding receipt дополнительно отличаются STAGE5_DEV.md и tests/test_stage5_sample_results.py; они не подменяют три основных source-input расхождения. Точный разбор — UZK_INPUT_BINDING_STOP.json.

Действия по контракту: STOP. Не восстановлены старые файлы, не обновлены receipts, не перегенерирован УЗК, не заменена принятая привязка новым расчетом и не объявлено, что изменения не влияют на результат. Чтобы продолжить, нужен отдельно согласованный актуальный checkpoint/reconciliation этих конкретных источников для выбранного SAMPLE_ONLY УЗК. Техническое совпадение выходного SHA не заменяет эту проверку.

АКЗ ACT SHA b640796570a8d04b6bca404205d99eae4c7005c700587812ed25bc99e715eb1a и APPLICATION SHA7766c92df82df2484eb82cd6917769f5d9ede912b3f0e8f26e3f8aaabdd357f2 совпали. Другие варианты не выбирались. Подробное сравнение фактического АКЗ с защищенным пунктом12 — NOT_RUN после STOP.

## 7. QUALITY DOCUMENT DYNAMIC FIELDS

Карта Этапа7 прочитана, ожидаемые9/2026, заказ1, масса15822,14кг, датакачества06.10.2026 и отдельные датыконтроля приняты как контракт. Product module/подстановка не реализованы из-за input gate STOP. Нет вывода с устаревшими/смешанными реквизитами.

## 8. POINT 12 IMMUTABILITY

Исходный template SHA проверен. Зафиксированная аудитом область T1 rows41–49/[2036,2702) не изменялась. Генерационный snapshot и сравнение output — NOT_RUN, поскольку output не создан. Не заявляется успешная проверка будущего форматирования.

## 9. POINT 13 GENERATED CONTENT

**NOT_RUN.** Новый документ/перечень не создан. Контракт остаётся: ВИК9/2026 и9/2026-1 от29.09.2026, УЗК9/2026 от01.10.2026, АКЗ9/2026 от05.10.2026, с явными TEST/SAMPLE_ONLY статусами.

## 10. POINT 14 SHORT LIST

**NOT_RUN.** Разрешён короткий список выбранных документов и постоянная фраза о хранении материалов/сертификатов. Исторический длинный список никуда не скопирован, экземпляры не назначались.

## 11. CHANGED FILES

Только новая папка `{OUT}`: preflight.py/finalize_stop.py, readonly JSON evidence и этот отчет. Stage7 product module, launcher, DOCX и selected TEST manifest не созданы. Полный перечень — CHANGED_FILES.json.

## 12. TARGETED VALIDATION

Выполнены preflight чтение DEV/канонического первого паспорта, template SHA, native visible VOM, сверка двух кандидатов Приложения№1, SHA УЗК/receipt/seal/input fingerprint binding, выбранной пары АКЗ/PAIR.json и её существующей source binding. **Gate: FAIL/STOP по трем input fingerprints УЗК.**

34 требуемых implementation/mutation/native acceptance сценария — NOT_RUN. Полная regression стадий1–6 не запускалась. Product BEFORE final tests / AFTER / CURRENT не выдаются за test proof: product tests не начинались, реализация отсутствует.

## 13. NATIVE WORD VALIDATION

**NOT_RUN.** Новый ВИК и документ качества не создавались; normal reopen,2страницы,protected12,Word numbering/fields/signatures проверять не на чем. Прежнее native audit evidence не выдается за E2E новой генерации.

## 14. OUTPUT DOCX PATH / SHA

**NOT_CREATED / NOT_APPLICABLE.** Официальный и TEST документ качества не создан, существующие Word не сохранены/не изменены.

## 15. EVIDENCE INTEGRITY

SOURCE_INTEGRITY.json: **{'PASS' if not changes else 'FAIL'}**, {len(before)} baseline source/product files, изменено{len(changes)}. Stage1–6 source files и DEV сохранены. Новая evidence manifest записана последней; старые audit/fix/sample receipts не переписаны. Production registry не открывался/не хэшировался.

## 16. REGRESSION IMPACT

Product changes: NONE. Только audit/preflight evidence. Нет изменений стадий1–6, ВОМ, справочника, Профилей, Word-шаблонов, исходных документов или сертификатов. Расхождения входов УЗК обнаружены до наших действий и не исправлялись.

## 17. NOT_RUN

Stage7 implementation/core/launcher, новый ВИК, генерация качества, frozen assembly manifest, новые УЗК/АКЗ/Приложение№1, targeted implementation tests, native Word output E2E, full regression1–6, production, macros, source saves, registry writes/sync, web, PDF, EXE, deploy, official issuance, Stage8.

## 18. READY FOR INDEPENDENT REVIEW

**NO** — реализации/выходного документа нет. Preflight blocker доступен для проверки по exact expected/actual SHA. Выполнено пользовательское условие STOP при несовпадении принятой input binding УЗК. После отчета STOP; Stage8 не начат.
'''
(OUT/'STAGE7_IMPLEMENTATION_STOP.md').write_text(report,encoding='utf8')
save('CHANGED_FILES.json',{'scope':str(OUT),'paths':[str(p) for p in sorted(OUT.iterdir()) if p.is_file()] + [str(OUT/'CHANGED_FILES.json'),str(OUT/'EVIDENCE_MANIFEST.json')],'product_code_changes':[]})
save('EVIDENCE_MANIFEST.json',{'state':'STOP_UZK_ACCEPTED_INPUT_BINDING_MISMATCH','ready_for_independent_review':False,'files':{str(p):sha(p) for p in sorted(OUT.iterdir()) if p.is_file() and p.name!='EVIDENCE_MANIFEST.json'},'written_last':True})
print(json.dumps({'state':'STOP','output_docx':None,'changed_source_files':changes,'input_mismatches':drift,'report':str(OUT/'STAGE7_IMPLEMENTATION_STOP.md')},ensure_ascii=False,indent=2))
