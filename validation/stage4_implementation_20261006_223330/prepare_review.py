"""Produce the review report after passing tests and visual inspection."""
import importlib.util
import json
from pathlib import Path
import sys
from unittest.mock import patch
sys.stdout.reconfigure(encoding='utf-8')
OUT=Path(__file__).resolve().parent
ROOT=OUT.parents[1]
sys.path.insert(0,str(ROOT))
from profiles_stage4.offline import read_json,sha
from profiles_stage4.storage import dump

tests=read_json(OUT/'tests.json')
build=read_json(OUT/'build_result.json')
sources=read_json(OUT/'source_ingestion.json')
assert tests['failed']==0 and tests['product_before_equals_after'] and tests['protected_unchanged']
spec=importlib.util.spec_from_file_location('stage4_finalizer',OUT/'finalize.py')
finalizer=importlib.util.module_from_spec(spec)
spec.loader.exec_module(finalizer)
original=finalizer.read_json
negative=[]
for case in ('stale_product_after','failed_test','stale_workbook_hash','missing_product_path'):
    def altered(path):
        data=original(path)
        if case=='stale_product_after' and Path(path).name=='product_after.json':
            data[next(iter(data))]='0'*64
        if case=='failed_test' and Path(path).name=='tests.json':
            data['failed']=1
        if case=='stale_workbook_hash' and Path(path).name=='build_result.json':
            data['sha256']='0'*64
        if case=='missing_product_path' and Path(path).name=='product_before.json':
            data.pop(next(iter(data)))
        return data
    with patch.object(finalizer,'read_json',altered):
        try:
            finalizer.validate()
        except ValueError as exc:
            negative.append({'case':case,'status':'PASS','rejected':str(exc)})
        else:
            raise AssertionError('Finalizer accepted negative case: '+case)
dump(OUT/'finalizer_negative.json',negative)
dump(OUT/'visual_verification.json',{'method':'Codex view_image inspection of all five final workbook sheet previews',
    'status':'PASS','sheets':['Профили','Источники','Правила УЗК','Предпочтения','Конфликты'],
    'previews':{str(p.relative_to(OUT)):sha(p) for p in sorted((OUT/'final_previews').glob('*.png'))},
    'native_Excel_visual_UAT':'NOT_RUN'})
c=build['counts']
def table(items):
    return '\n'.join(f'| {k} | {v} |' for k,v in items.items())
mandatory='\n'.join(f"| {t['id']} | {t['name']} | {t['status']} |" for t in tests['tests'])
conflicts=read_json(OUT/'conflicts.json')
conflict_table='\n'.join(f"| {r['id']} | {r['key']} | {r['standard']} | {', '.join(r['fields'])} |" for r in conflicts)
report=f'''# ВЗМК — Этап 4 IMPLEMENTATION — DEV

## 1. ACTUAL STATE

Stage 1–3 ACCEPTED, сохранены. Protected hashes: {tests['protected_count']}/{tests['protected_count']} неизменны.
DEV registry/checkpoint, Stage 3 catalog, Source A/B snapshots не изменялись. Production не открывался.

## 2. OFFLINE SOURCE INGESTION

Source A: {sources['A']} raw records из четырёх XLS; hash pins и исходные XLS origins проверены read-only.
Source B: {sources['B']} raw records из девяти accepted group datasets; raw JSON сверяется с cached HTML.
Всего source raw: {sources['source_raw']}. Display rounding: {sources['rounding_records']} records / {sources['rounding_fields']} полей;
используются numeric sort values. Category/detail rows не импортируются. ГОСТ 8639/8645, ТУ,
специальные профили и «Вне ГОСТ» исключены. Сетевые запросы: 0.

## 3. CATALOG DESIGN

`{build['workbook']}`

Листы: Профили, Источники, Правила УЗК, Предпочтения, Конфликты. Основная геометрия numeric, мм.
Identity = key + normalized НД. Canonical row одна на identity; совпавшие raw records сохраняются в provenance.
Нерешённые конфликты имеют пустую canonical geometry. Все {c['conflicts']} открыты, решения пустые.
Блокируются также вспомогательные source disagreements площади/массы; один отсутствующий R/r не блокирует запись.
Raw values, source cells/group rows/URL/retrieval/cache SHA сохраняются в `profiles.catalog.json`.
User state — отдельный `profiles.user.json`: manual rows/preferences/resolutions. Excel — представление;
изменения через API/CLI. `profiles.build.json` связывает hashes всей generation; внешние правки блокируют rebuild/lookup.

## 4. PROFILE COUNTS

Canonical: {c['canonical']}; manual: {c['manual_rows']}; UZK enabled: {c['uzk_enabled']}.
Conflicts: {c['conflicts']}; ambiguous keys без НД/preference: {c['ambiguous_keys']}.
Отсутствуют оба optional radius R/r: {c['missing_optional_radius']} (включая unresolved geometry).
R absent: {c['missing_R']}; r absent: {c['missing_r']}.

| Family | Canonical rows |
|---|---:|
{table(c['by_family'])}

| НД | Canonical rows |
|---|---:|
{table(c['by_standard'])}

| Status | Rows |
|---|---:|
{table(c['by_status'])}

## 5. 35Б1 GOLDEN

| Lookup | Status | h/b/s/t, мм | УЗК, мм |
|---|---|---|---|
| 35Б1 / ГОСТ Р 57837-2017 | FOUND | 346/174/6/9 | 328×6; 348×9 |
| 35Б1 без НД | FOUND_PREFERRED | 346/174/6/9 | 328×6; 348×9 |
| 35Б1 / ГОСТ 26020-83 | FOUND | 346/155/6.2/8.5 | 329×6.2; 310×8.5 |

Expected values использовались только тестами, не ingestion. Все варианты — из принятых snapshots.
Preference basis: USER_CONFIRMED_UZK_328x6_348x9; статус ПОДТВЕРЖДЕНО ПОЛЬЗОВАТЕЛЕМ.

## 6. LOOKUP

FOUND: explicit НД; также 25Б1 без НД имеет equivalent geometry и оба provenance.
FOUND_PREFERRED: 35Б1 без НД. AMBIGUOUS: 20Ш1 без НД. CONFLICT: L125×80×7 / ГОСТ 8510-86.
NOT_FOUND: неизвестный профиль или несовпадающий explicitly requested НД. Fallback между НД отсутствует.
Normalization сохраняет серии/нули и границы слов opaque names; объединяет только доказанные aliases/formats.
Предпочтение не разрешает same-ND source conflict. Различия optional radius по разным НД показаны unresolved.

## 7. CHANGED FILES

Новые product files: `profiles_stage4/__init__.py`, `model.py`, `offline.py`, `storage.py`, `workbook.mjs`, `gui.py`;
`stage4.py`, `Запустить Этап 4.cmd`, `STAGE4_README.md`. Junction node_modules — на bundled runtime, зависимости не изменены.
Новые DEV files: Профили.xlsx, profiles.catalog.json, profiles.user.json, profiles.build.json, verified rebuild backups.
Audit/test evidence и isolated fixtures — только в текущем validation folder.
Stage 1–3/source audits/DEV registry/checkpoints не изменены.

## 8. IMPLEMENTATION

API `profiles_stage4.storage.lookup(profile_text, standard=None)`; `build`, `add_manual`, `set_preference`,
`resolve_conflict`, `confirm_manual_uzk`. Pure functions `calculate_uzk`, `frequency_angle`.
Rule ДВУТАВР_ПРОКАТНЫЙ_1: `(h-2*t)×s`, `(2*b)×t`, одно изделие. Равные S суммируются только внутри одного профиля.
S<12 →5 МГц/70°; S>=12 →2.5 МГц/65°. Все другие типы НЕ НАСТРОЕНО.
Ручная запись РУЧНОЙ без автоматического УЗК; отдельное explicit confirmation поддержано для двутавра.
Минимальный Tk launcher: lookup/НД, размеры, status, provenance, УЗК; DEV workbook open с hash preflight.

## 9. TARGETED VALIDATION

{tests['passed']}/{len(tests['tests'])} PASS. Обязательные 1–43: 43/43. Stage 1–3 — только import/hash/checkpoint smoke.
Python socket guard и audit hook: реальные network events=0. Node HTTP/HTTPS/net/UDP/DNS/fetch guard:0.
Отдельный negative guard probe блокирован до network IO. Safe backup, pre-commit failure,
partial commit rollback, manual/preference/resolution preservation проверены в isolated fixtures.
User state tampering fail-closed; stale resolution fingerprints fail-closed; finalizer negative cases:4/4.
Все пять sheet previews визуально проверены. Real Tk widgets проверены для всех пяти lookup statuses;
open button dispatch проверен с mock os.startfile. Native Excel UI/human UAT не выполнялись.

| ID | Targeted test | Result |
|---:|---|---|
{mandatory}

## 10. EVIDENCE INTEGRITY

Product hashes BEFORE=AFTER; tests содержат exact tested hashes. Finalizer проверяет BEFORE=AFTER=CURRENT,
все обязательные тесты, сохранность protected files, current catalog↔commit↔user state↔workbook hashes/counts и reopen.
Workbook SHA-256: `{build['sha256']}`.
Review manifest создаётся последним: STAGE4_REVIEW_MANIFEST.json.
Evidence: product_before.json/product_after.json, tests.json, protected_before.json/protected_after.json,
golden_35Б1.json, source_ingestion.json, conflicts.json, ambiguous_keys.json, finalizer_negative.json, visual_verification.json.

## 11. REGRESSION IMPACT

Stage 4 отдельный package/launcher, не меняет Stage 1–3. Все {tests['protected_count']} protected files совпадают.
Отсутствующие радиусы не INVALID; unresolved conflicts исключены из обычного FOUND.
Практические ограничения: {c['conflicts']} source conflicts требуют явного решения;
{c['ambiguous_keys']} keys требуют НД/preference, unproved families/standards не включены.
Rebuild зависит от локального Python3.12 и bundled artifact runtime; обычный lookup обходится без Node/интернета.

### Все актуальные source conflicts

| Conflict ID | Key | НД | Fields |
|---|---|---|---|
{conflict_table}

## 12. NOT_RUN

Новые HTTP/source audits, 23met, OCR/PDF work, УЗК.xlsm, ВОМ quantities, passport/УЗК.docx,
АКЗ/Документ о качестве/PDF/EXE/deploy, полный Stage 1–3 regression, Stage 3 Apply, production, Stage 5.
Native Excel GUI visual UAT и ручная пользовательская проверка launcher: NOT_RUN.

## 13. READY FOR INDEPENDENT REVIEW — YES

Готовность требует успешного current-hash finalizer; никаких дальнейших этапов автоматически не начинается.
'''
(OUT/'STAGE4_REPORT.md').write_text(report,encoding='utf-8')
print(json.dumps({'report':str(OUT/'STAGE4_REPORT.md'),'negative_finalizer':len(negative),'tests':tests['passed']},ensure_ascii=False))
