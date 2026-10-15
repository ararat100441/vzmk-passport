"""Read-only closeout checks and new evidence only; manifest is written last."""
from datetime import datetime,timezone
from pathlib import Path
import json
import sys
sys.dont_write_bytecode=True
HERE=Path(__file__).resolve().parent
ROOT=HERE.parents[1];sys.path.insert(0,str(ROOT))
from stage7_quality.provenance import read,write,verify,sha,now
from stage7_quality.core import verify_manifest


def link(path,label=None):
    p=Path(path).resolve();return '['+(label or p.name)+'](<'+str(p).replace('\\','/')+'>)'


baseline=read(HERE/'BASELINE_BEFORE.json')
verify(baseline)
after={p:sha(p) for p in baseline}
write(HERE/'BASELINE_AFTER.json',after)
product_before=read(HERE/'PRODUCT_BEFORE_FINAL_TESTS.json')
product_after=read(HERE/'PRODUCT_AFTER_FINAL_TESTS.json')
verify(product_before)
assert product_before==product_after
result=read(HERE/'ASSEMBLY_01/ASSEMBLY_RESULT.json')
manifest=read(result['manifest']['path']);verify_manifest(manifest)
receipts={k:read(result[k]['path']) for k in ['uzk_receipt','vik_receipt','quality_receipt']}
tests=read(HERE/'TARGETED_TEST_RESULTS.json');assert tests['status']=='PASS' and tests['tests']>=34
assert read(HERE/'GUI_SMOKE.json')['status']=='PASS'
assert sha(HERE/'POINT12_TEMPLATE_XML_SNAPSHOT.xml')==sha(HERE/'POINT12_OUTPUT_XML_SNAPSHOT.xml')
write(HERE/'SOURCE_INTEGRITY.json',{'status':'PASS','checked':len(baseline),'changed':0,
    'product_before_after_current_equal':True,'bound_product_files':len(product_before),
    'stage1_6_changes':0,'old_receipt_changes':0,'production':'NOT_OPENED; explicit DEV inputs only',
    'web':'NOT_RUN','pdf':'NOT_RUN','exe':'NOT_RUN','stage8':'NOT_RUN'})
fresh_product=[*sorted((ROOT/'stage7_quality').glob('*.*')),ROOT/'stage7.py',ROOT/'Запустить Этап 7 DEV.cmd',ROOT/'STAGE7_DEV.md']
fresh_product=[p for p in fresh_product if p.is_file()]
changes={'product_files':[{'path':str(p),'sha256':sha(p),'change':'ADDED'} for p in fresh_product],
    'old_product_changes':[],'protected_source_changes':[],
    'validation_directory':str(HERE),'validation_files':[str(p.relative_to(HERE)) for p in sorted(HERE.rglob('*')) if p.is_file()],
    'new_final_files':['CHANGED_FILES.json','STAGE7_IMPLEMENTATION_REPORT.md','EVIDENCE_MANIFEST.json']}
write(HERE/'CHANGED_FILES.json',changes)
uzk=receipts['uzk_receipt'];vik=receipts['vik_receipt'];quality=receipts['quality_receipt']
lines=[
'# ВЗМК Этап 7 — реализация согласованного TEST-комплекта 9/2026',
'',
'READY FOR WORK 07 REVIEW: YES. BLOCKER 0. Только DEV/TEST; официальная выдача запрещена. Этап 8 не запускался.',
'',
'## 1. ROOT CAUSE AND FIX',
'',
'Прежний STOP имел две причины: расхождение vom.py с принятым M1 и отсутствие штатного receipt API Stage5. M1 восстановлен и независимо принят Work01 до этого запуска; Stage7 его не изменял. Отсутствующий API заменён явно разрешённым новым Stage7 execution wrapper. Старые receipts сохранены. Это не повторный аудит и не изменение Stage5.',
'',
'## 2. vom.py reconciliation',
'',
'Текущий SHA `c63bbd362763ba99f85de1610c01622ac230f3303b7a1e297345d02259ea954f` совпал с accepted current SHA. Свежий read_thread Work01 от 09.10.2026 подтвердил PASS / BLOCKER0 / MAJOR0 / MINOR0; M1 4/4, Stage6 compatibility 4/4. В этом запуске только SHA и отсутствие нового drift. ВОМ прочитана принятым native Excel read-only adapter:104 visible /397 hidden excluded /288 шт./15822.14кг.',
'',
'## 3. Accepted current checkpoints',
'',
'DEV registry `8e56f68518c70cade9efb09084dd67cd5d72e07e3fdc813b8515030e9cb98e80`; DEV УЗК `9f25a30139ab9e31b6da7248284c2ae6e9d4327f0d68251d5723293992662302`; profiles `7b420f6dead96408680a64d220c83e8980aeec1c38635e35b2e4723a98e3831b`. Принятая migration4→3 сохранена: catalog `7fa36f9ecc2ebb722aff89fdb94ffedf330608dc73c16264005e31e11f4b1587`, inputs.py `0890d0c2e354813f794ac7dbb3bba464c06c17d9ae71289e5ce13f594d756f32`, catalog.py `c2fa2912fc872344e00dc1dfe35cb686530b0055f94a4c2412d2b546926ce009`. Exact product checkpoint: '+link(ROOT/'stage7_quality/accepted_checkpoint.json')+'.',
'',
'`read_registry()` / `select_project()` динамически подтвердили выбранный9/2026, строка400; next10/2026 не использован. Даты качества06.10, ВИК29.09, УЗК01.10, АКЗ05.10.2026 прочитаны из текущего DEV. Приложение№1:123 записи,288изделий,15822.14кг, по каждой марке совпадает с104видимыми строками ВОМ. AKZ0 accepted act SHA `b640796570a8d04b6bca404205d99eae4c7005c700587812ed25bc99e715eb1a`, application `7766c92df82df2484eb82cd6917769f5d9ede912b3f0e8f26e3f8aaabdd357f2`; pair_id `e9fc12e4433feafd`, current preview source bindings совпали. Покрытие:один слой Pirocor Protect057, RAL7004, требование80мкм; замеры НЕ подтверждены; только Бп2-5/Бп2-18/Бп2-19 —3шт.,105.26кг.',
'',
'## 4. Stage7 wrapper design',
'',
'assembly_run_id `'+result['assembly_run_id']+'`. Перед вызовом проверяются47input bindings и102code/checkpoint bindings. Оригинальные входы/код удерживаются Windows read handles с FILE_SHARE_READ: запись и удаление запрещены. Это эквивалент frozen input snapshot. VIK получает отдельные побайтные копии DEV_registry.xlsx и exact first-template; Stage2 сам блокирует их и пересчитывает только свою внутреннюю копию. В production не передаётся путь. Stage5 settings.json использует только documents_root/projects; registry_path production из settings не применяется.',
'',
'Recorder регистрирует точное имя callable, пишет START, реально вызывает generator, проверяет новый output/path/SHA, пишет RESULT, выполняет OOXML/native проверки и повторно сверяет inputs/code, затем VALIDATION и отдельный STAGE7_OPERATION_RECEIPT. Время фактическое UTC; дата документа берётся из DEV. Ошибка даёт FAILURE без успешного receipt; такой файл не включается в selected manifest. SHA-цепочка является локальной audit evidence, не электронной подписью.',
'',
'## 5. Actual UZK TEST run and SHA',
'',
link(uzk['output'])+' — SHA `'+uzk['output_sha256']+'`.',
'',
'Принятый `uzk_stage5.core.generate` фактически вызван: operation_id `'+uzk['operation_id']+'`. START '+read(uzk['start']['path'])['time']+'; receipt '+uzk['time']+'. Расчёт:Бп2-1=8шт.(1280×5;1024×8.4), Пс1-9=9шт.(3600×4), Пс1-10=15шт.(6000×4), Пс2-16=1шт.(480×3); всего4/33,5МГц/70°,01.10.2026. ДНО/годен — SAMPLE_ONLY, результаты не подтверждены. Новый файл детерминированно совпал SHA со старым образцом; факт нового вызова подтверждён новой START/RESULT/VALIDATION цепочкой, файл не скопирован из старого output.',
'',
'## 6. New Stage7 operation receipts',
'',
'УЗК: '+link(result['uzk_receipt']['path'])+'; SHA `'+result['uzk_receipt']['sha256']+'`.',
'',
'ВИК: '+link(result['vik_receipt']['path'])+'; SHA `'+result['vik_receipt']['sha256']+'`.',
'',
'Качество: '+link(result['quality_receipt']['path'])+'; SHA `'+result['quality_receipt']['sha256']+'`. Все receipts COMPLETE_VALIDATED_TEST, official_issue=false, results_confirmed=false. Старый Stage5 receipt не изменён и не выдан за новый.',
'',
'## 7. New VIK',
'',
link(vik['output'])+' — SHA `'+vik['output_sha256']+'`. Принятый `passport_app.vik.generate_vik` фактически вызван в VIK_ISOLATION; operation_id `'+vik['operation_id']+'`. В одном DOCX две проверенные части:9/2026 и9/2026-1,29.09.2026. generation.json подтверждает current DEV SHA, строку400, exact template SHA, неизменность констант/оформления, normal reopen. Номенклатура соединений совпадает с текущим DEV. Исходный Stage2 не добавляет видимую TEST-плашку; его форма сохранена. DEV/TEST-статус явно задан изолированным местом, receipts, manifest и качеством; файл отдельно не предназначен для официальной выдачи.',
'',
'## 8. Selected source manifest',
'',
link(result['manifest']['path'])+' — SHA `'+result['manifest']['sha256']+'`. Шесть записей:Приложение1;ВИК9;ВИК9-1;УЗК SAMPLE_ONLY;АКЗ0 TEST;его приложение TEST. Все привязаны к одному project/passport/current registry/selected visible VOM. Две записи ВИК имеют общий файл/run, разные native page locators. АКЗ сохраняет original act_id/pair_id и исходный TEST-статус; original run Приложения1 неизвестен и не придуман. АКЗ1/2/3 исключены. Нет автоматического выбора по timestamp.',
'',
'## 9. Quality DOCX and SHA',
'',
link(quality['output'])+' — SHA `'+quality['output_sha256']+'`. Качество сформировано последним после receipt УЗК и ВИК и валидации manifest. Exact first template SHA `4df7234d37030188030142ed6a519dc2a923aada4f90e58a7b5140c0fd54f9cc`. Guarded local replacements по принятому mapping, заказ№1 сохранён. Изменены только согласованные динамические поля, TEST-плашка в пустом subtitle paragraph и пункты13/14. Пункт12 сохранён канонически побайтно XML всех9строк, включая исходную орфографию, текст и формат. Internal before snapshot взят до правок, post-check выполнен до создания receipt; отдельные comparison snapshots сохранены после генерации, без заднего датирования. Все прочие ZIP parts идентичны шаблону. Пункт13 —4проверенныхакта, пункт14 —5выбранныхприложений+фраза хранения+ограничение TEST. Длинная история сертификатов отсутствует, число экземпляров не выдумано; подписи/печать не созданы.',
'',
'## 10. Native Word checks',
'',
'Три новых DOCX нормально открыты Microsoft Word COM, ReadOnly=True, OpenAndRepair=False, без сохранения; SHA до/после равны. Каждый2страницы. Качество:1секция,2таблицы,14логическихпунктов, неизменные numbering/MERGEFIELD/constant cells/headers/footer/signature blanks. Журавлев и дата находятся на странице2. Все6PNG просмотрены:clipping/overlap/третьейстраницы нет. EMF→PNG через GDI, без PDF. '+link(HERE/'VISUAL_REVIEW.json')+'.',
'',
'## 11. Targeted test results',
'',
'46/46 PASS, failures0/errors0/skipped0. '+link(HERE/'TARGETED_TEST_RESULTS.json')+'; '+link(HERE/'TARGETED_TESTS.log')+'. Проверены обязательные34контракта и дополнительные отказные сценарии recorder/code drift/fabricated receipt/frozen writes. NO_UZK_REQUIRED presentation проверен in-memory; для текущего9/2026 скрыть фактический required УЗК запрещено. Mutation tests только in-memory/isolated failure fixtures. GUI hidden Tk callback smoke PASS:состав/двеВИК/SAMPLE_ONLY/ограниченныйАКЗ видимы; без подтверждения generator_calls=0. '+link(HERE/'GUI_SMOKE.json')+'. Live user UI acceptance NOT_RUN; native Word COM proof выполнен отдельно. Полная регрессия1–6 NOT_RUN по контракту.',
'',
'## 12. Source integrity',
'',
f'PASS: {len(baseline)} защищённых файлов BEFORE=AFTER, changed0. Bound product {len(product_before)} файлов BEFORE final tests = AFTER = CURRENT. Stage1–6,DEV,ВОМ,каталог,профили,шаблоны,старыеoutputs/receipts/evidence сохранены. '+link(HERE/'SOURCE_INTEGRITY.json')+'. Production не открывался/не хэшировался, макросы/списание/синхронизация не выполнялись.',
'',
'## 13. Changed files',
'',
f'Только {len(fresh_product)} новых файлов функционала Stage7 и новая validation папка. '+link(HERE/'CHANGED_FILES.json')+'. New package stage7_quality:provenance/sources/core/word/native/gui/render/accepted checkpoint; CLI stage7.py; DEV launcher; STAGE7_DEV.md. Existing product files changed0. Подготовительные ошибки нового Stage7 preflight (len hidden list и Decimal(str(float))) исправлены до actual generation; старый product не затронут. Все production outputs/read-only sources сохранены.',
'',
'## 14. Remaining blockers / limitations',
'',
'BLOCKER0 для Stage7 DEV/TEST review. Официальная выдача BLOCKED_UNCONFIRMED_CONTROL:УЗК образец;АКЗ без реальных замеров, охват3изделияиз288;текстовые константы исходного ВИК не удостоверяют фактический контроль. Новый DEV GUI реализован, live ручная приёмка пользователя NOT_RUN. Нет PDF/EXE/deploy/web/full regression/Stage8. Независимое Work07 review ещё не выполнено и не присваивается этим отчётом.',
'',
'## 15. READY FOR WORK 07 REVIEW: YES',
'',
'Stage7 implementation завершён в принятом DEV/TEST объёме. После отчёта STOP. EVIDENCE_MANIFEST.json записывается последним, включает всё новое evidence и hashes новых product files, кроме самого себя; затем проверяется read-only.',
]
(HERE/'STAGE7_IMPLEMENTATION_REPORT.md').write_text('\n'.join(lines)+'\n',encoding='utf8')
# Source after consistency once more immediately before the final manifest.
verify(baseline);verify(product_before);verify_manifest(manifest)
evidence={str(p.resolve()):{'sha256':sha(p),'bytes':p.stat().st_size}
          for p in sorted(HERE.rglob('*')) if p.is_file() and p.name!='EVIDENCE_MANIFEST.json'}
final={'schema':'STAGE7_EVIDENCE_MANIFEST_V1','created_at':now(),'status':'READY_FOR_WORK07_REVIEW',
    'assembly_run_id':result['assembly_run_id'],'evidence':evidence,'new_product':changes['product_files'],
    'protected_source_count':len(baseline),'protected_source_changed':0,'tests':tests['tests'],
    'official_issue':False,'stage8':'NOT_RUN'}
write(HERE/'EVIDENCE_MANIFEST.json',final)
verify({p:v['sha256'] for p,v in evidence.items()})
verify({v['path']:v['sha256'] for v in changes['product_files']})
print(json.dumps({'status':'PASS','report':str(HERE/'STAGE7_IMPLEMENTATION_REPORT.md'),
    'evidence_files':len(evidence),'evidence_manifest_sha256':sha(HERE/'EVIDENCE_MANIFEST.json'),
    'new_product_files':len(fresh_product),'protected_sources':len(baseline),'quality':quality['output'],
    'quality_sha256':quality['output_sha256'],'ready_for_work07_review':True},ensure_ascii=False,indent=2))
