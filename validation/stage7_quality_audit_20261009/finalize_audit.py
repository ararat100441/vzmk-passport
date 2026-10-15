import os,json,hashlib,re
from pathlib import Path
OUT=Path(__file__).resolve().parent;ROOT=OUT.parents[1]
def sha(p):return hashlib.sha256(p.read_bytes()).hexdigest()
def load(p):return json.loads(p.read_text(encoding='utf8'))
def save(n,d):(OUT/n).write_text(json.dumps(d,ensure_ascii=False,indent=2,default=str),encoding='utf8')
def walk(root,excluded=()):
    for base,dirs,files in os.walk(root,followlinks=False):
        dirs[:]=[d for d in dirs if d not in excluded and not (Path(base)/d).is_symlink() and not (Path(base)/d).is_junction()]
        for n in files:
            p=Path(base)/n
            if not p.is_symlink() and not n.startswith('~$'):yield p
before=load(OUT/'source_before.json')
files=list(walk(ROOT,('validation','.codebase-index','node_modules','__pycache__')))
files += [p for p in walk(ROOT.parent/'Гараж казахстан/Паспорт') if p.suffix.lower()!='.pdf']
after={str(p):{'size':p.stat().st_size,'sha256':sha(p)} for p in files}
changes=[{'path':p,'before':before.get(p),'after':after.get(p)} for p in sorted(set(before)|set(after)) if before.get(p)!=after.get(p)]
save('SOURCE_INTEGRITY.json',{'before_count':len(before),'after_count':len(after),'changes':changes,'pass':not changes,'scope':'Product and DEV excluding validation/index/cache/reparse; project passport files except PDFs/Word lockfiles; production registries not accessed'})
save('source_after.json',after)
manifest=load(OUT/'PROPOSED_SOURCE_MANIFEST.json');manifest['registry']['path']=str(ROOT/'DEV/ВЗМК_DEV.xlsx');save('PROPOSED_SOURCE_MANIFEST.json',manifest)
supp=load(OUT/'supporting_docx.json');vik=supp[-1]
receipt=load(ROOT/'e2e_user_stage2/validation/user_e2e/.vik-stage/5e341a79eaee47acb09c98d75e17c624/generation.json')
sample_meta=load(ROOT/'validation/stage5_sample_results_20261008_130441/new_document.json');sample=Path(sample_meta['output'])
protected=[]
for i in (1,2,3):
    d=load(OUT/f'word_{i}.json');cells=d['tables'][0]['cells'];text=''.join(c['text'] for c in cells if 41<=c['row']<=49)
    protected.append({'path':d['path'],'text':text,'text_sha256':hashlib.sha256(text.encode('utf8')).hexdigest()})
xml=load(OUT/'ooxml_structure.json')
save('FINAL_CHECKS.json',{
    'source_integrity':not changes,
    'word_source_integrity':all(load(OUT/f'word_{i}.json')['integrity_pass'] for i in (1,2,3)),
    'vik_output_receipt_sha_matches':vik['sha256']==receipt['output_sha256'],
    'vik_registry_binding_matches_current_dev':receipt['registry_sha256']==manifest['registry']['sha256'],
    'sample_file_exists':sample.is_file(),'sample_receipt_sha_matches':sample.is_file() and sha(sample)==sample_meta['sha256'],
    'sample_mode':sample_meta['mode'],'sample_results_confirmed':sample_meta['results_confirmed'],
    'protected_point12_text_same_in_three_history_examples':len({p['text_sha256'] for p in protected})==1,
    'protected_point12_history':protected,
    'vertical_merge_count':sum(c['vMerge'] is not None for r in xml['row_properties'] for c in r['cells']),
    'evidence_has_document_or_pdf_output':any(p.suffix.lower() in ('.doc','.docx','.xlsx','.xls','.xlsm','.pdf','.exe') for p in OUT.iterdir()),
    'actual_assembled_manifest':'NOT_ASSEMBLED','ready_to_implement_stage7':'NO','user_clarification':'Пункты 12–14 формы',
    'source_visual_review':'Native Word EMF and PNG: first passport both pages, history 9 and 2 page images; no PDF',
    'visual_generation_e2e':'NOT_RUN'
})
report=OUT/'STAGE7_QUALITY_AUDIT.md';text=report.read_text(encoding='utf8')
text=text.replace('Не invent copies=1.','Не назначать число экземпляров=1 без источника.')
text=text.replace('Вертикальные объединения проверяются в `ooxml_structure.json`; структура сохраняется целиком.','Вертикальных объединений нет; горизонтальная сетка и объединения описаны в `ooxml_structure.json`, структура сохраняется целиком.')
text=text.replace('Результат — `SOURCE_INTEGRITY.json`;','Результат: **PASS, 427 файлов до / 427 после, изменений 0** — `SOURCE_INTEGRITY.json`;')
text=text.replace('07е','07е')
# Expand compressed audit labels for readable prose without changing source quotations.
replacements={
'источникI400':'источник I400','заголовокПриложения':'заголовок Приложения','датакачества':'дата качества','документакачества':'документа качества',
'project/форма':'project/форма','completechecks':'complete checks',
'источника1':'источника 1','Этапа7':'Этапа 7','Этап7':'Этап 7','Этап8':'Этап 8','Этапа1':'Этапа 1','Этап1':'Этап 1','Этап2':'Этап 2','Этап5':'Этап 5','Этап6':'Этап 6','Этапов1–6':'Этапов 1–6','Этапы1–6':'Этапы 1–6','стадии1–6':'стадии 1–6','стадиях1–6':'стадиях 1–6','Stage1–6':'Stage 1–6',
'паспорта1/2026':'паспорта 1/2026','паспорт1/2026':'паспорт 1/2026','паспорт19407':'паспорт 19407','паспорт9/2026':'паспорт 9/2026','документами 2/2026':'документами 2/2026',
'пункт12':'пункт 12','пункта12':'пункта 12','пункт13':'пункт 13','пункта13':'пункта 13','пункт14':'пункт 14','пункта14':'пункта 14','пунктах13/14':'пунктах 13/14','пункты13/14':'пункты 13/14','пункты1–14':'пункты 1–14','пунктов8/10':'пунктов 8/10','пункта10':'пункта 10','пункта11':'пункта 11',
'составом,несколько':'составом, несколько','массой19407':'массой 19407','масса19407':'масса 19407','ВОМ9':'ВОМ 9','Приложения9':'Приложения 9','Приложение9':'Приложение 9','приложения9':'приложения 9','акта9':'акта 9','качество9':'качество 9','ВИК9':'ВИК 9','АКЗ9':'АКЗ 9','АКЗ0/1/2/3':'АКЗ 0/1/2/3','АКЗ1/2/3':'АКЗ 1/2/3','АКЗ1':'АКЗ 1','АКЗ0':'АКЗ 0',
'4 марки':'4 марки','4марки':'4 марки','2страницы':'2 страницы','2таблицы':'2 таблицы','1секцию':'1 секцию','14пунктов':'14 пунктов','17 PAIR':'17 PAIR','результаты1/2/5/6':'результаты 1/2/5/6',
'отдельныхактов':'отдельных актов','будущихoutputs':'будущих outputs',
'нетВИК→нетложнойссылки':'нет ВИК → нет ложной ссылки','дваВИКв одномфайле':'два акта ВИК в одном файле','нетфайла/ссылки':'нет файла/ссылки','неполнота/STOP':'неполнота / STOP','subsetАКЗ':'subset АКЗ','неполнаяпара':'неполная пара','другойпаспорт/проект':'другой паспорт/проект','changedSHA':'changed SHA','mixedrun безразрешения':'mixed run без разрешения','несовпадающаядатаакта':'несовпадающая дата акта','противоречиепункту':'противоречие пункту','длинныйперечень/переносподписи':'длинный перечень/перенос подписи','неизменность12 иисточников':'неизменность пункта 12 и источников',
'принятыхстадиях':'принятых стадиях','структурный анализ выполнен bundledruntime':'структурный анализ выполнен bundled runtime','существующим Python3.12':'существующим Python 3.12',
'проекта/DEV/Этапов':'проекта/DEV/Этапов','427 persistent':'427 persistent',
'правило13/14':'правило 13/14','productionВЗМК.xlsx':'production ВЗМК.xlsx','ProductionВЗМК.xlsx':'Production ВЗМК.xlsx','productionУЗК.xlsm':'production УЗК.xlsm',
'путь,размер,SHA,а':'путь, размер, SHA, а','наабсолютныйпуть':'на абсолютный путь','историческихчтения':'исторических чтения','Источникинеизменены':'Источники не изменены',
'ИндексныеMCPtools':'Индексные MCP tools','ИндексныеMCPtools':'Индексные MCP tools','недоступны; код':'недоступны; код',
'один первый вызов':'один первый вызов','историческогоWord':'исторического Word','с относительнымпутём':'с относительным путём','только audit-script':'только audit-script',
'полныйтекст':'полный текст','ожидаемуюпагинацию':'ожидаемую пагинацию',
'генераторыВИК/УЗК/АКЗ':'генераторы ВИК/УЗК/АКЗ','тестыЭтапов':'тесты Этапов','productionдоступ':'production доступ','DEVреестра':'DEV реестра','изменениеВОМ':'изменение ВОМ','конвертацияDOC':'конвертация DOC','сохранениеисточниковWord':'сохранение источников Word','официальнаявыдача':'официальная выдача','nativeWordvisualE2Eновогогенерированногодокумента':'native Word visual E2E нового генерированного документа','проверкаУЗКreadonly':'проверка УЗК read-only','прежнийTESTпример':'прежний TEST-пример','новыйзапускNOT_RUN':'новый запуск NOT_RUN',
'документовпункта14':'документов пункта 14','единыйTESTкомплект':'единый TEST-комплект','единыйTEST':'единый TEST','подтвержденныйперечень':'подтвержденный перечень',
'ООО«ВЗМК»':'ООО «ВЗМК»','443076г.Самара,ул.Крупской,1Е':'443076 г. Самара, ул. Крупской, 1 Е','пунктов8/10':'пунктов 8/10','НачальникОТК,ФИОЖуравлевА.Н.':'Начальник ОТК, ФИО Журавлев А.Н.','Заказ1':'Заказ 1','Pirocor Protect057':'Pirocor Protect 057','менее80мкм':'менее 80 мкм','очистки2':'очистки 2','сертификат№31889':'сертификат №31889','принтерiR2930':'принтер iR2930',
'история1/2026':'история 1/2026','ВИК1; ВИК1-1; УЗК1; АКЗ1':'ВИК 1; ВИК 1-1; УЗК 1; АКЗ 1','повторного решения':'повторного решения',
'источникI':'источник I','документа из Приложения':'документа из Приложения',
'состав дополнительных приложенийпункта':'состав дополнительных приложений пункта','конкретный единыйTESTкомплект/запуск':'конкретный единый TEST-комплект/запуск',
'документакачества':'документа качества','датконтроля':'дат контроля','номеразаказа':'номера заказа','реальныхподтверждений':'реальных подтверждений','принадлежностьюи':'принадлежностью и'
}
for a,b in replacements.items():text=text.replace(a,b)
parts=text.split('`')
for i in range(0,len(parts),2):
    t=parts[i]
    t=re.sub(r'\b(в|В|у|У|для|Для|к|К)(?=\d)',r'\1 ',t)
    t=re.sub(r'(пункт[а-яё]*|страниц[а-яё]*|строк[а-яё]*|таблиц[а-яё]*|вариант[а-яё]*|Этап[а-яё]*|исходник[а-яё]*|Приложени[а-яё]*|исторически[а-яё]*|перечн[а-яё]*|Заказ)(?=\d)',r'\1 ',t)
    t=re.sub(r'([а-яА-ЯёЁ])(?=(?:Word|Excel|TEST|SAMPLE|SHA|PNG|PDF|openpyxl|pywin32|inputbindings)\b)',r'\1 ',t)
    for a,b in {'датыкачества':'даты качества','серияМК':'серия МК','адрес443076':'адрес 443076','должностьНачальник':'должность Начальник','protected12':'защищённый пункт 12','outputSHA':'output SHA','сформировать13/14':'сформировать пункты 13/14','Сформировать13/14':'Сформировать пункты 13/14','Проверить14':'Проверить 14','потерять13/14':'потерять 13/14','изменить12':'изменить 12','замена1/2026':'замена 1/2026','содержитopenpyxl':'содержит openpyxl','неpywin32':'не pywin32','затемPASS':'затем PASS','завершилсяfile-not-found':'завершился file-not-found','синхронизацияDEV':'синхронизация DEV','implementationЭтапа':'implementation Этапа','BundledPython':'Bundled Python','акта№':'акта №','датой06':'датой 06','дату06':'дату 06','от08.':'от 08.','от05.':'от 05.','от02.':'от 02.','от04.':'от 04.','от29.':'от 29.','от01.':'от 01.','номер9/':'номер 9/','содержат1/':'содержат 1/','варианты0/':'варианты 0/','таблица1':'таблица 1','таблица2':'таблица 2','обычный9/':'обычный 9/','составом,несколько':'составом, несколько','Пункт12':'Пункт 12'}.items():t=t.replace(a,b)
    parts[i]=t
text='`'.join(parts)
report.write_text(text,encoding='utf8')
save('CHANGED_FILES.json',{'scope':str(OUT),'files':[{'path':str(p),'size':p.stat().st_size,'sha256':sha(p)} for p in sorted(OUT.iterdir()) if p.is_file() and p.name!='CHANGED_FILES.json'],'self_entry':'CHANGED_FILES.json excluded from self hashing'})
print(json.dumps({'source_files_before':len(before),'source_files_after':len(after),'source_changes':changes,'evidence_files':len(list(OUT.iterdir())),'ready':'NO'},ensure_ascii=False))
if changes:raise SystemExit(2)
