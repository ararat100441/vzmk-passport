"""Record the actual first-passport input blocker; do not alter product or workbooks."""
from pathlib import Path
import json,hashlib,os
OUT=Path(__file__).resolve().parent;ROOT=OUT.parents[1]
def read(p):return json.loads(Path(p).read_text(encoding='utf8'))
def sha(p):return hashlib.sha256(Path(p).read_bytes()).hexdigest()
def save(p,d):
    with Path(p).open('x',encoding='utf8') as f:json.dump(d,f,ensure_ascii=False,indent=2,default=str);f.flush();os.fsync(f.fileno())
accepted=read(OUT/'ACCEPTED_BEFORE.json');stage9=read(OUT/'SOURCE_BEFORE.json');seed=read(OUT/'SEED_BEFORE.json')
changes=[p for mapping in (accepted,stage9,seed['files']) for p,h in mapping.items() if sha(p)!=h]
current=read(OUT/'CURRENT_PROJECT_TEMPLATES.json');changes+=[r['path'] for r in current['files'] if sha(r['path'])!=r['sha256']]
assert not changes,changes
baseline=read(ROOT/'validation/stage9_gui_099_20261012_01/BINDING_AFTER.json')['build']
exe=Path(baseline['exe']);build={str(p.relative_to(exe.parent)):sha(p) for p in exe.parent.rglob('*') if p.is_file()}
assert build==baseline['files'] and sha(exe)==baseline['exe_sha256']
checks=[]
for n in range(4):
    act=next(r for r in current['files'] if r['variant']==n and r['kind']=='act')
    application=next(r for r in current['files'] if r['variant']==n and r['kind']=='application')
    cells=act['cells'];valid=n==0
    checks.append({'variant':n,'source_binding_status':'PASS_FIRST_PASSPORT_PAIR' if valid else 'FAIL_OTHER_PROJECT_AND_NONFIRST_APPLICATION',
      'act':act['path'],'act_sha256':act['sha256'],'application':application['path'],'application_sha256':application['sha256'],
      'act_cells':{k:cells[k] for k in ('E4','H4','E5','E6','E7','B44')},'application_A1':application['cells']['A1'],
      'generation_test':'NOT_RUN_INPUT_GATE','native_excel_test':'NOT_RUN_INPUT_GATE','automatic_ranges':'NOT_IMPLEMENTED_INPUT_GATE'})
save(OUT/'FIRST_PASSPORT_GATE.json',{'project':'Гараж Казахстан','first_passport':'1/2026','status':'STOP','reason':'Variants 1-3 do not satisfy the new same-project first-passport template requirement. Do not adapt or relabel them as approved project templates.','checks':checks})
save(OUT/'SOURCE_INTEGRITY_AFTER.json',{'status':'PASS','accepted_stage1_8_count':len(accepted),'accepted_stage1_8':accepted,'stage9_files':stage9,'sealed_dev_unchanged':True,'current_project_templates_unchanged':True,'changes':changes,'exe_version':'0.9.9','exe':str(exe),'exe_sha256':sha(exe),'onedir_count':len(build),'onedir_unchanged':True,'product_files_modified':0,'production_writes':0})
report='''**ВЗМК — автоматическое заполнение АКЗ / проверка обязательных входов / STOP**

13.10.2026, Europe/Samara. Продолжение существующего проекта. Новый модуль автоматического заполнения, настройки диапазонов и новая EXE **не реализованы**: обязательный источник для трёх вариантов не соответствует прямому ограничению текущего запроса. Это блокер входных шаблонов, не ошибка Office worker и не основание менять принятый Stage 6.

**Требование пользователя:** брать соответствующие пары АКЗ №0/1/2/3 только из первого паспорта проекта, не использовать шаблоны других проектов. Ранее принятый генератор позволял заменять исторические реквизиты предоставленных форм. Такое прошлое принятие не подтверждает соответствие новому требованию использовать именно первый паспорт текущего проекта.

Первый паспорт установлен принятым `read_registry()` / `select_project()` только по sealed DEV реестру: «Гараж Казахстан», **1/2026**, строка **353**. 09.09.2026 в Project.first_date — дата окончания; акт №0 датирован 08.09.2026, это отдельная дата АКЗ, не смешивается с датой окончания. После этого прочитаны восемь exact известных файлов первого паспортного каталога. Реальные реестры не открывались. Прочитаны только шаблоны, без Excel COM/макросов/сохранения.

| Вариант | Акт `4!E4`, `4!E6`, `4!E7` | Приложение `A1` | Gate |
|---|---|---|---|
| №0 | 1/2026; ООО «СТК ВК»; объект парковки | АКЗ 1/2026 от 08.09.2026 | PASS: допустимая пара первого паспорта |
| №1 | 1/Б2; ООО «РУССОКО»; 2-ЭГ-2019-800/1-Б2-КМД | документ о качестве 4/2026 от 24.09.2026 | FAIL: другой проект + не первый паспорт приложения |
| №2 | 1/2025; ООО «Велесстрой»; ВС.24.ПО-Э1.2-РД-ЭЦ-КМД | документ о качестве 5/2026 от 25.09.2026 | FAIL: другой проект + не первый паспорт приложения |
| №3 | 1/2025; ООО «Велесстрой»; ВС.24.ПО-Э1.2-РД-ЭЦ-КМД | документ о качестве 6/2026 от 30.09.2026 | FAIL: другой проект + не первый паспорт приложения |

Текущие файлы и sealed DEV копии **8/8 SHA-identical**. Следовательно, причина не в устаревшей DEV копии. Путь папки сам по себе не доказывает принадлежность её содержимого проекту. Exact filenames, paths, SHA, значения и формулы — CURRENT_PROJECT_TEMPLATES.json / TEMPLATE_BINDINGS_BEFORE.json / FIRST_PASSPORT_GATE.json. Ни одну чужую пару не использовали для нового заполнения и не переназывали правильным шаблоном.

У №0 исходные Pirocor Protect 057, RAL 7004, 1 слой, 80 мкм подтверждены. У чужих форм №1–3 материалы/толщины были лишь прочитаны как диагностика: они не становятся нормативами «Гараж Казахстан». В частности, `∑1+2=180` нельзя преобразовать в самостоятельно придуманную норму второго слоя или перенести из другого проекта.

**Что реализовано:** только bounded проверка обязательных источников и SHA evidence; новых функций продукта **0**. Существующий Stage 6 код/EXE не изменён. Не заявляем частичную №0 генерацию за выполнение требования проверить все четыре варианта.

**Настройки диапазонов:** NOT_IMPLEMENTED_INPUT_GATE. Значения температуры/влажности/толщины не придуманы и не сохранены в реестры/шаблоны. Существующий ClimateDay поддерживает запись на дату и конфликт STOP, однако настраиваемые диапазоны/автомоделирование ещё не реализованы. Согласованная формула точки росы не получена; текущий Stage 6 прямо помечает её неутверждённой. Запрошен текст/коэффициенты/источник. До согласования Td не рассчитывать; TEST моделирование не выдавать за фактические измерения.

**Результаты:** проверка идентичности источников 8/8 PASS; first-passport binding №0 PASS / №1–3 FAIL. Native Excel, генерация XLSX, тесты диапазонов, формул/оформления новых результатов, EXE integration новой функции — **NOT_RUN_INPUT_GATE**. Это не targeted implementation PASS. Никаких новых XLSX не создано.

**Source integrity:** 56 принятых Stage 1–8 файлов unchanged, Stage 9 unchanged, sealed DEV source unchanged, восемь прочитанных текущих шаблонов unchanged. Прежняя EXE 0.9.9 SHA `21ff2f7cc0d4fca746094c46d3594ae0b1dbf991c40266195e22c1030c0ccadc`; все 1117 onedir файлов unchanged. Изменены/добавлены только validation/akz_auto_20261013_01 scripts/evidence/report. Реальных списаний, официального закрытия, production writes, GUI/Computer Use и deployment нет. Office-процессы не создавались и не завершались.

**Для продолжения нужны:** точные правильные пары АКЗ №1/2/3 и приложения первого паспорта каждого подключаемого проекта; по текущему стенду — именно 1/2026 «Гараж Казахстан». Нужны оригинальные формы с правильной принадлежностью и нормативами, а не разрешение произвольно переименовать чужие книги. Если правильные пары уже существуют в другой папке — указать exact paths. Для автоматического Td отдельно требуется согласованная формула; остальные климатические TEST поля можно реализовать с явно нерассчитанным Td после решения источников.

**READY FOR WORK REVIEW: NO.** Реализация и требуемые positive tests всех четырёх вариантов не выполнены. **Официальный выпуск: НЕ ГОТОВ.** Задание на возобновление подготовлено, независимый review не запускался. После отчёта **STOP** на доказанном input blocker.
'''
with (OUT/'REPORT.md').open('x',encoding='utf8') as f:f.write(report)
assignment='''**АКЗ AUTO — blocked implementation handoff**

Не запускать review самостоятельно. Кандидат не готов: READY FOR WORK REVIEW: NO.
Прочитать REPORT.md / FIRST_PASSPORT_GATE.json / CURRENT_PROJECT_TEMPLATES.json.
Сначала получить от пользователя exact правильные пары первого паспорта; текущие №1–3 содержат чужие проекты и приложения 4/5/6. Не менять Stage1–8 и не делать silent substitution/adaptation. Не считать нахождение файла в папке проекта достаточным доказательством.
После решения входов: новый Stage9 adapter/settings отдельно от шаблонов/реестров; один TEST климатический набор на календарную дату с durable conflict STOP; редактируемые диапазоны на вариант/слой, нормативы только своей утверждённой пары; никакого выдуманного standalone80 для cumulative180. Td только согласованная формула, иначе явное NOT_CALCULATED. Сохранять формулы/мерджи/стили/материалы/цвета/слои/нормативы/реквизиты; марки/числа только проверенные видимые VOM rows; blank != 0. Все output TEST и official STOP.
Новые positive native Excel tests №0/1/2/3, common-day/restart/conflict/invalid-range/SHA/cross-project/official STOP/transaction tests, structure/formulas/styles/visual verification на своих DEV копиях. Не выдавать старые Stage6/Stage9 тесты за новый функционал. Full evidence/hash maps и changed files; независимый verdict лишь после отдельного разрешённого запуска Work. STOP после отчёта.
'''
with (OUT/'RESUME_ASSIGNMENT.md').open('x',encoding='utf8') as f:f.write(assignment)
save(OUT/'RESULT.json',{'status':'STOP_INPUT_CONFLICT','implementation':'NOT_RUN','product_changes':0,'valid_first_passport_pairs':[0],'invalid_first_passport_pairs':[1,2,3],'native_excel':'NOT_RUN','new_exe':'NOT_BUILT','ready_for_work_review':False,'official_issue_ready':False,'stop':True})
files={str(p):sha(p) for p in OUT.rglob('*') if p.is_file()}
save(OUT/'EVIDENCE_MANIFEST.json',{'schema':'AKZ_AUTO_FIRST_PASSPORT_INPUT_STOP_V1','status':'STOP_INPUT_CONFLICT','files':files,'source_integrity':'PASS','ready_for_work_review':False})
bad=[p for p,h in files.items() if sha(p)!=h];assert not bad
save(OUT/'SEAL_VERIFY.json',{'status':'PASS','manifest_sha256':sha(OUT/'EVIDENCE_MANIFEST.json'),'file_count':len(files),'mismatches':bad})
print(json.dumps({'status':'STOP_INPUT_CONFLICT','accepted_unchanged':len(accepted),'onedir_unchanged':len(build),'manifest_sha256':sha(OUT/'EVIDENCE_MANIFEST.json'),'new_product_code':0,'ready_for_work_review':False},indent=2))
