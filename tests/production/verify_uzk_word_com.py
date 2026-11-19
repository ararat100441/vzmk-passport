r"""PR14 Word Gate A: force 3+ UZK pages and check header repeats.

Reads the corrected VZMK UZK DOCX for SHA and structural preflight. Uses
ONLY a temporary copy and explicitly synthetic welds. Requires a NEW
--output diagnostic directory to inspect all PDF pages visually.

  py -3.12 -B tests/production/verify_uzk_word_com.py ^
      --template "C:\...\ВЗМК_УЗК_ШАБЛОН_на_утверждение.docx" ^
      --output "C:\...\PR14-new-uzk-word-diagnostics"
"""
import argparse
import hashlib
import json
import re
import shutil
import sys
import tempfile
import zipfile
from pathlib import Path
from xml.etree import ElementTree as E

from pypdf import PdfReader

ROOT=Path(__file__).resolve().parents[2]
sys.path.insert(0,str(ROOT))
from passport_app.vik import word_session,open_word
from production_app.office import word_table
from production_app.real_forms import UZK_HEADERS

TOKEN='{{UZK_DETAILS_TABLE}}'
NS='{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'


def sha(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def verify_marker(path):
    with zipfile.ZipFile(path) as package:
        body=E.fromstring(package.read('word/document.xml')).find(NS+'body')
    if body is None:
        raise ValueError('Нет основного текста Word.')
    blocks=list(body)
    matches=[]
    for index,node in enumerate(blocks):
        if node.tag!=NS+'p':
            continue
        text=''.join(x.text or '' for x in node.iter(NS+'t')).strip()
        if TOKEN in text:
            matches.append((index,text))
    if len(matches)!=1 or matches[0][1]!=TOKEN:
        raise ValueError('Нужен один отдельный абзац-маркер УЗК.')
    index=matches[0][0]
    if index<1 or blocks[index-1].tag!=NS+'tbl':
        raise ValueError('Проверяемая форма не содержит прежнюю таблицу '
                         'реквизитов непосредственно перед УЗК.')
    return {'one_standalone_marker':True,'requisites_before_uzk':True}


def synthetic_weld_rows(count=60):
    # Complete table with 8 confirmed TEST rows, not a fake actual UZK report.
    return [
        ('TEST-UZK-'+str(i).zfill(3),
         'Сварное соединение № '+str(i)+', длина 540 мм, толщина 8 мм',
         'Тестовый сварщик, клеймо TEST', '19.10.2026',
         '5', '70', 'TEST: дефекты отсутствуют','ТЕСТОВЫЕ ДАННЫЕ')
        for i in range(1,count+1)
    ]


def verify_saved_table(doc,original_count):
    if doc.Tables.Count!=original_count+1:
        raise RuntimeError('Word снова присоединил таблицу УЗК к реквизитам.')
    found=[]
    for index in range(1,doc.Tables.Count+1):
        table=doc.Tables(index)
        try:
            if table.Rows(1).Cells.Count!=8:
                continue
            headings=[table.Cell(1,j).Range.Text.strip('\r\x07 ')
                      for j in range(1,9)]
            if headings==list(UZK_HEADERS[:8]):
                found.append(table)
        except Exception:
            continue
    if len(found)!=1:
        raise RuntimeError('Не найдена ровно одна самостоятельная таблица УЗК с 8 колонками.')
    table=found[0]
    if table.Rows.AllowBreakAcrossPages:
        raise RuntimeError('Word разрешает разрыв записи соединения между страницами.')
    if not bool(table.Rows(1).HeadingFormat):
        raise RuntimeError('В Word не сохранён повторяемый заголовок строк УЗК.')
    doc.Repaginate()
    pages=doc.ComputeStatistics(2)
    if not isinstance(pages,int) or pages<3:
        raise RuntimeError('Word не сформировал три страницы для проверки повтора шапки.')
    return {'pages':pages,
            'tables':doc.Tables.Count,
            'uzk_columns':8,
            'uzk_rows_unsplit':True,
            'uzk_header_repeat_property':True}


def check_pdf(pdf_file):
    reader=PdfReader(str(pdf_file),strict=False)
    if len(reader.pages)<3:
        raise RuntimeError('Нативный Word PDF не содержит 3 страниц.')
    report=[]
    for index,page in enumerate(reader.pages,1):
        if float(page.mediabox.width)<400 or float(page.mediabox.height)<400:
            raise RuntimeError('Неожиданная геометрия PDF страницы '+str(index))
        try: text=page.extract_text() or ''
        except Exception:text=''
        norm=''.join(text.casefold().split())
        report.append({'page':index,
                       'has_uzk_mark':bool(re.search('test-uzk-[0-9]{3}',norm)),
                       'has_uzk_header':'марка' in norm,
                       'text_extracted':bool(norm)})
    # Require repeated header on pages where a synthetic weld is actually
    # present, if the PDF's text layer is readable. Images require human check.
    if report[0]['text_extracted']:
        for data in report[1:]:
            if data['has_uzk_mark'] and not data['has_uzk_header']:
                raise RuntimeError('Word PDF не содержит повторяемой шапки '
                                   'УЗК на странице '+str(data['page']))
    return report


def main(template_path,output):
    original=Path(template_path).resolve(strict=True)
    if original.suffix.lower()!='.docx':
        raise ValueError('Укажите исправленный Word DOCX УЗК.')
    before=sha(original)
    structure=verify_marker(original)
    destination=Path(output).resolve()
    destination.mkdir(parents=True,exist_ok=False)
    # Keep only explicitly requested diagnostics. The user's source stays
    # untouched; the edited Word document is a COPY inside a new folder.
    scratch=destination/'synthetic-uzk.docx'
    pdf=destination/'synthetic-uzk.pdf'
    shutil.copyfile(original,scratch)
    with word_session() as app:
        doc=open_word(app,scratch,readonly=False)
        try:
            original_count=doc.Tables.Count
            rows=synthetic_weld_rows()
            table=word_table(doc,TOKEN,UZK_HEADERS[:8],rows)
            if table.Rows(1).Cells.Count!=8:
                raise RuntimeError('Word не создал восемь ячеек таблицы.')
            doc.Save()
        finally:
            doc.Close(SaveChanges=False)

        proof=open_word(app,scratch,readonly=True)
        try:
            doc_data=verify_saved_table(proof,original_count)
            if TOKEN in proof.Content.Text:
                raise RuntimeError('В документе остался маркер таблицы.')
            proof.ExportAsFixedFormat(
                OutputFileName=str(pdf),
                ExportFormat=17,OpenAfterExport=False,
                OptimizeFor=0,Range=0,
            )
        finally:
            proof.Close(SaveChanges=False)

    pdf_data=check_pdf(pdf)
    if sha(original)!=before:
        raise RuntimeError('Исходный исправленный Word шаблон изменился!')
    result={
        'status':'PENDING_VISUAL_REVIEW',
        'technical_checks_passed':True,
        'human_review_required':True,
        'visual_rule':'Просмотреть все страницы Word/PDF: 8 видимых колонок, '
                      'повтор заголовка на 2-й/3-й, без обрезки.',
        'source_sha256_unchanged':True,
        'template':str(original),
        'diagnostic_word':str(scratch),
        'diagnostic_pdf':str(pdf),
        **structure, **doc_data,
        'pdf_page_checks':pdf_data,
    }
    (destination/'WORD_REVIEW.json').write_text(
        json.dumps(result,ensure_ascii=False,indent=2),encoding='utf-8')
    return result


if __name__=='__main__':
    parser=argparse.ArgumentParser()
    parser.add_argument('--template',required=True)
    parser.add_argument('--output',required=True,help='НОВАЯ папка для Word/PDF диагностики')
    args=parser.parse_args()
    try:
        print(json.dumps(main(args.template,args.output),
                         ensure_ascii=False,indent=2))
    except BaseException as exc:
        print('FAIL '+type(exc).__name__+': '+str(exc),file=sys.stderr)
        raise
