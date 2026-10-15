"""Write the final report, then the manifest as the last filesystem mutation."""
from pathlib import Path
from datetime import datetime, timezone, timedelta
import hashlib
import json

OUT=Path(__file__).resolve().parent
ROOT=OUT.parents[1]
MANIFEST=OUT/'REVIEW_MANIFEST.json'
def sha(p):return hashlib.sha256(Path(p).read_bytes()).hexdigest()
def read(n):return json.loads((OUT/n).read_text(encoding='utf-8-sig'))
def link(p):return '['+Path(p).name+'](<'+Path(p).as_posix()+'>)'
def dump(p,d):p.write_text(json.dumps(d,ensure_ascii=False,indent=2),encoding='utf-8')
def require(v,m):
    if not v:raise AssertionError(m)
require(not MANIFEST.exists(),'MANIFEST_ALREADY_EXISTS_STOP')
r=read('resume_readonly_validation.json')
require(r['status']=='PASS','RESUME_NOT_PASS_STOP')
sources=read('sources_after.json')
products=read('current_product_hashes.json')
tested=read('tested_product_after.json')
for p,h in sources.items():require(sha(p)==h,'SOURCE_CHANGED_BEFORE_MANIFEST_STOP: '+p)
for p,h in products.items():require(sha(p)==h,'PRODUCT_CHANGED_BEFORE_MANIFEST_STOP: '+p)
for x in r['six_outputs']:require(sha(x['path'])==x['sha256'],'OUTPUT_CHANGED_BEFORE_MANIFEST_STOP')
tz=timezone(timedelta(hours=4))
def mtime(n):return datetime.fromtimestamp((OUT/n).stat().st_mtime,tz).isoformat()
lines=[]
def add(text=''):lines.append(text)
add('# ВЗМК — ЭТАП 6 / Part 2 — RESUME AFTER ESCAPE')
add('\nREADY FOR INDEPENDENT REVIEW — YES. Финализация существующего TEST-результата; официальная выдача STOP.\n')
add('## 1. ACTUAL STATE AFTER ESCAPE\n')
add('На первом чтении после resume уже существовали финальный продукт, три FINAL_TEST пары (6 XLSX и 3 PAIR.json), полный журнал 14 targeted PASS, 33 unit PASS и совпадающие tested-product BEFORE/AFTER. До этой финализации итоговый отчёт, sources_after, current product hashes и final REVIEW_MANIFEST отсутствовали. Точное время нажатия Escape не записано; нельзя достоверно отнести завершение фонового прогона к моменту до или после нажатия. Подтверждено фактическое состояние на resume, без такого предположения.\n')
add('| Существовавшее доказательство | Время изменения, UTC+04 |')
add('|---|---|')
for n in ('real_generation.json','tested_product_before.json','targeted_validation.json','unit_tests.txt','tested_product_after.json'):add(f'| {n} | {mtime(n)} |')
add('\nСтарые initial-пары, native_real_outputs.json, initial GUI events/screenshots и attempt1/2/3 сохранены как история. Они не заменяют проверку финальных шести файлов. Ранние неуспешные прогоны не скрыты: attempt1 — слишком строгое float-сравнение в R1 harness; attempt3 — SAVED_INITIAL_VIEW_STOP при ошибочном размещении проверки. На входе resume исправления уже существовали; полный финальный прогон и текущие SHA подтверждают их конечное состояние. В этой финализации дефектов не обнаружено и исправлений не выполнено.\n')
add('## 2. PRODUCT FILES\n')
add('Part 2: добавлены part2.py и test_stage6_part2.py; расширены gui.py и STAGE6_DEV.md. Остальные принятые файлы Part 1 сохранены. CURRENT SHA-256:\n')
add('| Exact product path | SHA-256 |')
add('|---|---|')
for p,h in sorted(products.items()):add(f'| {link(p)} | `{h}` |')
add('\n## 3. COMPLETED VALIDATION\n')
add('Повторно использованы 14 проверок из финального targeted_validation.json, привязанных к tested BEFORE=AFTER=CURRENT, и 33/33 unit-теста. На resume дополнительно выполнены только чтение OOXML/cached values точных шести XLSX и сверка SHA исходных, продуктовых файлов и доказательств. Ни генерация, ни targeted harness повторно не запускались.\n')
add('| Существующая targeted проверка | Результат |')
add('|---|---|')
for x in r['targeted_checks_reused']:add(f"| {x['name']} | {x['status']} |")
add('\nCapacity: №1 61 — PASS / 62 — CAPACITY_STOP; №2 47 — PASS / 48 — CAPACITY_STOP; №3 94 — PASS / 95 — CAPACITY_STOP. Locking: реальный Windows share=0 handle, primary error сохранён при cleanup failure, unknown/replaced pending не удалён, восстановление в fresh process и retry — PASS. Все данные этих тестов находятся в sandbox_final2, отдельном от финальных шести XLSX.\n')
add('## 4. NOT_COMPLETED VALIDATION\n')
add('Обязательных незавершённых проверок в пределах RESUME-контракта не осталось. Финальная визуальная GUI-проверка после последнего исправления saved initial view не доказана: имеющиеся screenshots относятся к initial версии. Она не подменяется PASS. Сохранённый initial view A1 и native Excel проверки финальных файлов подтверждены; GUI не запускался на resume согласно прямому запрету пользователя.\n')
add('## 5. SIX TEST XLSX PATHS / SHA\n')
add('Все шесть CURRENT SHA совпадают с PAIR.json и финальным native targeted result; до/после текущего read-only чтения SHA неизменны.\n')
add('| Вариант / файл | Exact path | SHA-256 |')
add('|---|---|---|')
for x in r['six_outputs']:add(f"| АКЗ-{x['variant']} / {x['kind']} | {link(x['path'])}<br>`{x['path']}` | `{x['sha256']}` |")
for v,mark,mass,area,total in ((1,'Бп2-5','37.20','1.412',66),(2,'Бп2-18','31.32','1.184',52),(3,'Бп2-19','36.74','1.393',99)):
    add(f'\n## {v+5}. АКЗ-{v} RESULT\n')
    add(f'PASS: акт 9/2026-{v}; дата 05.10.2026; приложение с тем же номером, датой, объектом и ACT-ID. Строка 5: {mark}, Балка, 1 шт., {mass} кг, {area} м²; удельные и общие значения совпадают при количестве 1. Итоговая строка {total} видима: 1 шт., {mass} кг, {area} м²; SUM сохранены, cached values совпадают. Старые строки очищены и скрыты, TEST/ЧЕРНОВИК присутствует в документе и печатном header. Старые номера, контрагенты, договоры и подписи из шаблонов отсутствуют; текущие реквизиты совпадают с принятой DEV preview.\n')
    if v==1:add('Слой: Norecoat FD Primer Grey, серый, требование 150 мкм.\n')
    if v==2:add('Слои: Триокор Мастик 4500 / Серый / 100 мкм; Триокор Финиш 5500 / 7015 / буквальное ∑1+2=180. Индивидуальная требуемая толщина второго слоя NOT_PROVEN/null; 80 мкм не выведены и не записаны.\n')
    if v==3:add('Слои: Триокор Мастик 4500 / Серый / 100 мкм; Триокор Мастик 4500 / Серый / 80 мкм; Триокор Финиш 5500 / 7015 / 60 мкм. Это требования шаблона, не измеренные результаты.\n')
add('## 9. CLIMATE / HISTORICAL SAFETY\n')
add('PASS: во всех слоях строки 30–34 (дата/поверхность/воздух/RH/Td) — НЕТ ДАННЫХ / ТЕСТ. Климатический receipt один и тот же день 2026-10-05, все четыре измерения null, confirmation=MISSING. Строки 26–28 измеренной толщины очищены. Серии, даты нанесения, время, инструменты, подписи и прочие фактические поля также MISSING; степени/способ не подтверждены (?). Во всех XML шести файлов отсутствует RANDBETWEEN; в актах остались только формулы ссылок номера/даты G8/I8, в приложениях только три итоговые SUM. L:O очищено. У АКЗ-2 Лист1 veryHidden, содержит только служебную TEST-метку, не имеет измерений/формул и не включён в обычную печать. ∑1+2=180 сохранено. Официальная выдача STOP (receipt, тексты документов, existing official-mode tests).\n')
add('## 10. NATIVE EXCEL RESULTS\n')
add('PASS на exact final six: normal load, ReadOnly, full calculation, Close(SaveChanges=False), SHA до=после. В каждой паре акт 1 страница и приложение 1 страница; service_printed=false. Проверены номера/даты/реквизиты/слои/исходные строки/итоги и отсутствие возврата исторических фактов после пересчёта. Основание: финальный targeted_validation.json, первый result, и соответствующий validate.py:39–49; product BEFORE/AFTER совпадает с CURRENT. Native COM повторно НЕ запускался: существующих привязанных доказательств достаточно. Физическая печать не выполнялась.\n')
add('## 11. №0 / R1 / R2 REGRESSION\n')
add('№0 core.py SHA неизменён относительно принятого independent R2 review: eede991cf2d2ca259f251ef9f79beb49542ea1e536ff2eea6f18cf9c23cd935d. Сохранены __init__, stage6.py, launcher.cmd и исходный test_stage6.py. GUI расширен для групп, поэтому его новый SHA проверяется финальным прогоном, а не старым №0 verdict. R1 native N3/N17/N18 — PASS; R2 реальные lock/cleanup/fresh-process recovery/retry — PASS; original primary error через actual Tk button handler — PASS. Принятые прежние №0 TEST outputs не перезаписаны; их SHA также входят в сохранённый source guard. Старый independent review подтверждает принятую исходную точку; текущий GUI подтверждён новым финальным targeted прогоном.\n')
add('## 12. SOURCE INTEGRITY\n')
add('Сверены все 534 пути sources_before.json: 532 неизменны; ровно два предусмотренных изменения Part 2 — gui.py и STAGE6_DEV.md. Неожиданных изменений нет. Stage 1–5 (main/stage2–5/passport_app/profiles_stage4/uzk_stage5, прежние tests, manifests и DEV outputs из снимка) сохранены. Шесть исходных АКЗ-1/-2/-3 шаблонов, исходные №0 шаблоны, DEV registries/checkpoints, реальный Справочник марок.xlsx, выбранный реальный ВОМ и исходные профили неизменны. Созданный ранее Part 2 DEV fixture также равен SHA из DEV_assignments.json; на resume назначения не менялись. Полный перечень exact paths/SHA — sources_before.json и sources_after.json.\n')
critical=[(p,h) for p,h in sources.items() if Path(p).name in ('ВЗМК_DEV.xlsx','УЗК_DEV.xlsm','Справочник марок.xlsx','АКЗ-1.xlsx','АКЗ-2.xlsx','АКЗ-3.xlsx','Приложение к АКЗ-1.xlsx','Приложение к АКЗ-2.xlsx','Приложение к АКЗ-3.xlsx')]
add('| Защищённый файл | SHA-256 BEFORE = AFTER |')
add('|---|---|')
for p,h in critical:add(f'| {link(p)} | `{h}` |')
add('\nProduction не использовался: этот resume выполнял только локальные чтения конкретных DEV/output/source путей и записи evidence. В сохранённом финальном harness network_events=[] и forbidden_production_events=[]; это Python audit-hook evidence, не системный сетевой/COM trace. Production registry не открывался для повторной проверки; доказательством не объявляется отсутствующий OS-wide аудит.\n')
add('## 13. EVIDENCE INTEGRITY\n')
add('CURRENT product hashes зафиксированы в current_product_hashes.json; для шести tested файлов TESTED_BEFORE = TESTED_AFTER = CURRENT. Точные шесть output hashes совпадают с финальным тестом и receipt. Исторические attempt/initial результаты не используются как proof нового кода. Финальный манифест создаётся после этого отчёта и всех JSON; содержит SHA всех существующих Part 2 evidence files, current product, 534 источников, six outputs, receipts и reused proof bindings. Манифест не хеширует сам себя. После его записи выполняется только read-only проверка перечисленных файлов/инвентаря и manifest-last.\n')
add('## 14. CHANGED FILES\n')
add('В IMPLEMENTATION Part 2 ранее: добавлены akz_stage6/part2.py и tests/test_stage6_part2.py; изменены akz_stage6/gui.py и STAGE6_DEV.md; созданы DEV fixture, final TEST-пары и изолированные validation artifacts. Эти изменения уже существовали на входе resume. В ЭТОМ RESUME продукт, XLSX и DEV/source файлы не изменялись. Добавлены только в существующем evidence каталоге:\n')
for n in ('resume_readonly.py','finalize_resume_evidence.py','sources_after.json','current_product_hashes.json','resume_readonly_validation.json','FINAL_REPORT.md','REVIEW_MANIFEST.json'):add('- '+str(OUT/n))
add('\n## 15. NOT_RUN\n')
add('На resume: генерация XLSX, Stage 6 code edits, новые DEV assignments, Computer Use, GUI, повторный COM, повторный targeted harness, полный Stage 1–5 regression, большой повторный аудит — NOT_RUN. Финальное ручное GUI/визуальное одобрение новой версии, физическая печать, OS-wide packet/process trace — NOT_RUN. Official issuance, production и Stage 7 — NOT_RUN / STOP. Эти пункты не представлены как PASS.\n')
add('## 16. READY FOR INDEPENDENT REVIEW — YES\n')
add('Обязательные доказательства RESUME-контракта подтверждены и связаны с текущими неизменёнными файлами. Результат готов для независимого review, не для официальной выдачи. После отчёта STOP.\n')
(OUT/'FINAL_REPORT.md').write_text('\n'.join(lines),encoding='utf-8')

# All artifact writes must precede the following manifest creation.
evidence={}
for p in sorted(OUT.rglob('*')):
    if p.is_file() and p!=MANIFEST:
        st=p.stat();evidence[str(p)]={'sha256':sha(p),'bytes':st.st_size,'mtime_ns':st.st_mtime_ns}
receipts={str(Path(x['path']).parent/'PAIR.json'):sha(Path(x['path']).parent/'PAIR.json') for x in r['six_outputs']}
manifest={
    'schema':'STAGE6_PART2_RESUME_REVIEW_MANIFEST_V1',
    'created_utc':datetime.now(timezone.utc).isoformat(),
    'ready_for_independent_review':'YES','official':'STOP',
    'operation':'Evidence-only finalization of existing result after Escape',
    'product_current_sha256':products,
    'tested_before_sha256':read('tested_product_before.json'),
    'tested_after_sha256':tested,
    'product_binding':'TESTED_BEFORE = TESTED_AFTER = CURRENT for all six tested files; other three accepted files unchanged',
    'source_before_sha256':read('sources_before.json'),
    'source_after_sha256':sources,
    'authorized_part2_changes':r['authorized_part2_changes'],
    'unexpected_source_changes':[],
    'six_outputs':r['six_outputs'],'receipts_sha256':receipts,
    'fixture':r['fixture'],
    'reused_final_evidence':{str(OUT/n):sha(OUT/n) for n in ('validate.py','targeted_validation.json','tested_product_before.json','tested_product_after.json','unit_tests.txt','real_generation.json')},
    'validation':{'targeted_pass':14,'unit_pass':33,'resume_readonly':'PASS','native_exact_six':'PASS_REUSED_HASH_BOUND','reruns':[]},
    'historical_only':['native_real_outputs.json','initial_real_generation.json','initial_native_GUI_events.json','attempt1/2/3 evidence','initial screenshots'],
    'not_run':['Regeneration','Code edits','New DEV assignments','Computer Use','GUI on resume','COM on resume','Test reruns','Full Stage 1-5 regression','Final manual GUI visual approval','Physical print','OS-wide network/process trace','Official issuance','Production','Stage 7'],
    'network_production_evidence':r['network_production_audit'],
    'network_production_evidence_limit':r['audit_limit'],
    'resume_changed_files':[str(OUT/n) for n in ('resume_readonly.py','finalize_resume_evidence.py','sources_after.json','current_product_hashes.json','resume_readonly_validation.json','FINAL_REPORT.md','REVIEW_MANIFEST.json')],
    'evidence_files':evidence,
    'manifest_last_rule':'This manifest is the last artifact write. Only read-only verification may follow. Self hash excluded.',
}
dump(MANIFEST,manifest)  # LAST WRITE
print(json.dumps({'ready':'YES','manifest':str(MANIFEST),'manifest_sha256':sha(MANIFEST),'evidence_files':len(evidence),'report':str(OUT/'FINAL_REPORT.md')},ensure_ascii=False))
