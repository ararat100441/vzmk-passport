r"""Native Word Gate A: independent UZK table in actual corrected VZMK form.

Usage:
  py -3.12 -B tests/production/verify_uzk_word_com.py ^
      --template "C:\...\ВЗМК_УЗК_ШАБЛОН_на_утверждение.docx"

Original DOCX is never opened by Word for editing. A unique temporary copy
is used, with SYNTHETIC UZK rows only. Word is always closed normally.
"""
import argparse
import hashlib
import json
import shutil
import sys
import tempfile
import zipfile
from pathlib import Path
from xml.etree import ElementTree as E

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
    """Require the anchor in its own body paragraph after requisites table."""
    with zipfile.ZipFile(path) as z:
        body=E.fromstring(z.read('word/document.xml')).find(NS+'body')
    children=list(body)
    matches=[]
    for index,element in enumerate(children):
        if element.tag!=NS+'p':
            continue
        text=''.join(t.text or '' for t in element.iter(NS+'t')).strip()
        if TOKEN in text:
            matches.append((index,text))
    if len(matches)!=1 or matches[0][1]!=TOKEN:
        raise ValueError('Требуется один маркер УЗК в отдельном абзаце.')
    position=matches[0][0]
    if position<1 or children[position-1].tag!=NS+'tbl':
        raise ValueError('Не воспроизведён исходный риск: предыдущая таблица реквизитов.')
    return {'anchor_isolated':True,
            'original_requisites_table_before_marker':True}


def native_validate(doc,initial_tables,header):
    if doc.Tables.Count != initial_tables+1:
        raise RuntimeError('Word объединил УЗК с другой таблицей.')
    found=[]
    for i in range(1,doc.Tables.Count+1):
        table=doc.Tables(i)
        try:
            width=table.Rows(1).Cells.Count
        except Exception:
            continue
        if width!=len(header):
            continue
        heading=[table.Cell(1,j).Range.Text.strip('\r\x07 ')
                 for j in range(1,len(header)+1)]
        if heading==list(header):
            found.append(table)
    if len(found)!=1:
        raise RuntimeError('После сохранения Word не найдена ровно одна 8-колоночная таблица УЗК.')
    doc.Repaginate()
    pages=doc.ComputeStatistics(2)
    if type(pages) is not int or pages<1:
        raise RuntimeError('Word не подтвердил страницы УЗК.')
    return {'tables':doc.Tables.Count,'uzk_columns':len(header),
            'pages':pages,'header_verified':True}


def main(path):
    template=Path(path).resolve(strict=True)
    if template.suffix.lower()!='.docx':
        raise ValueError('Нужен исправленный DOCX-шаблон подробного УЗК.')
    original=sha(template)
    structure=verify_marker(template)
    with tempfile.TemporaryDirectory(prefix='vzmk-uzk-native-word-') as tmp:
        scratch=Path(tmp)/'synthetic-uzk.docx'
        shutil.copyfile(template,scratch)
        with word_session() as app:
            doc=open_word(app,scratch,readonly=False)
            try:
                original_count=doc.Tables.Count
                synthetic_rows=[
                    ('SYNTHETIC-M1','L=250x8 mm','Test Welder',
                     '19.10.2026','5','70','NONE','ГОДЕН')
                ]
                table=word_table(doc,TOKEN,UZK_HEADERS[:8],synthetic_rows)
                if table.Rows(1).Cells.Count!=8:
                    raise RuntimeError('В таблице УЗК не 8 отдельных ячеек.')
                doc.Save()
            finally:
                doc.Close(SaveChanges=False)
            proof=open_word(app,scratch,readonly=True)
            try:
                report=native_validate(proof,original_count,UZK_HEADERS[:8])
                if TOKEN in proof.Content.Text:
                    raise RuntimeError('Маркер таблицы УЗК не удалён.')
            finally:
                proof.Close(SaveChanges=False)
    if sha(template)!=original:
        raise RuntimeError('Оригинальный Word-файл изменён: STOP.')
    return {'status':'PASS','scope':'TEMPORARY_SYNTHETIC_WORD_UZK',
            'source_sha256_unchanged':True,**structure,**report}


if __name__=='__main__':
    parser=argparse.ArgumentParser()
    parser.add_argument('--template',required=True)
    arguments=parser.parse_args()
    try:
        print(json.dumps(main(arguments.template),ensure_ascii=False,indent=2))
    except BaseException as exc:
        print('FAIL '+type(exc).__name__+': '+str(exc),file=sys.stderr)
        raise
