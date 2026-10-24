r"""PR14 Windows Excel: full multipage appendix readability and PDF audit.

Only source-copy operations. For visual review, use --output pointing to a
NEW, NONEXISTENT folder. The tool returns PENDING_VISUAL_REVIEW, never
declares final visual PASS simply because Excel emitted several PDF pages.

Example:
  py -3.12 -B tests/production/verify_appendix_print_com.py ^
    --appendix-one "C:\...\1_1_2026 - Приложение № 1.xlsx" ^
    --akz-appendix "C:\...\approved-blank-AKZ-appendix.xlsx" ^
    --output "C:\...\PR14-new-diagnostics"
"""
import argparse
import hashlib
import json
import re
import shutil
import sys
import tempfile
from pathlib import Path

from pypdf import PdfReader

ROOT=Path(__file__).resolve().parents[2]
sys.path.insert(0,str(ROOT))
from passport_app.excel import excel_session
from production_app.office import (
    fit_appendix_body, set_appendix_print_titles, verify_appendix_print_layout,
)


def sha(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def sample_rows(count=110):
    """Cover short, medium and realistically long text in one printer test."""
    result=[]
    for number in range(1,count+1):
        suffix=str(number).zfill(3)
        mark='SYNTHETIC-'+suffix
        if number%11==0:
            name=('Тестовое сборочное изделие с дополнительным обозначением '
                  'металлоконструкции и протоколом контроля '+suffix)
        elif number%5==0:
            name='Колонна сборная, тестовый элемент '+suffix
        else:
            name='Балка, учебный образец '+suffix
        result.append((mark,name,1,float(number),float(number),
                       float(number)/10,float(number)/10))
    return result


def normalize_pdf_text(value):
    """Comparable across PDF layout engines with artificial line breaks."""
    return ''.join((value or '').casefold().split())


def audit_pdf_pages(pdf_file,first_mark,last_mark):
    reader=PdfReader(str(pdf_file),strict=False)
    count=len(reader.pages)
    if not 3<=count<=20:
        raise RuntimeError(
            'Нужна многостраничная PDF-диагностика (включая 3-ю страницу); '
            'получено страниц: '+str(count)
        )

    records=[]
    for number,page in enumerate(reader.pages,1):
        width=float(page.mediabox.width)
        height=float(page.mediabox.height)
        if not 400<=width<=1300 or not 400<=height<=1400:
            raise RuntimeError('Неверный размер PDF страницы '+str(number))
        try:
            raw=page.extract_text() or ''
        except Exception:
            raw=''
        normalized=normalize_pdf_text(raw)
        markers={
            'has_mark_header': 'отправочнаямарка' in normalized,
            'has_name_header': 'наименованиемарки' in normalized,
            'has_quantity_header': 'количество' in normalized,
            'has_synthetic_rows': 'synthetic-' in normalized,
        }
        records.append({
            'page':number,
            'width_pt':round(width,2),
            'height_pt':round(height,2),
            'text_extracted':bool(normalized),
            **markers,
        })
    # Extraction may be inconclusive for specific fonts, but if it works on
    # the first page, missing headings on page 3 are a demonstrable FAIL.
    first=records[0]
    if first['has_mark_header'] and first['has_name_header']:
        for page in records[1:]:
            if not (page['has_mark_header'] and page['has_name_header']):
                raise RuntimeError(
                    'В PDF отсутствует повторяемый заголовок на странице '
                    +str(page['page'])+'; проверьте третью страницу.'
                )
    if first['text_extracted'] and any(not p['text_extracted'] for p in records):
        raise RuntimeError('Одна из страниц PDF не содержит извлекаемого текста.')
    return records


def export_copy(app, source, folder, label):
    original=sha(source)
    scratch=folder/(label+'.xlsx')
    if scratch.exists():
        raise FileExistsError('Тестовая книга уже существует: '+str(scratch))
    shutil.copyfile(source,scratch)
    result=folder/(label+'.pdf')
    if result.exists():
        raise FileExistsError('Диагностический PDF уже существует: '+str(result))

    synthetic=sample_rows()
    first=5
    last=first+len(synthetic)-1
    book=app.Workbooks.Open(
        Filename=str(scratch),ReadOnly=False,UpdateLinks=0,AddToMru=False,
        IgnoreReadOnlyRecommended=True,
    )
    try:
        if book.ReadOnly or book.Worksheets.Count!=1:
            raise RuntimeError('Неожиданная защищённая Excel книга: '+str(scratch))
        sheet=book.Worksheets(1)
        if 'отправочная марка' not in str(sheet.Cells(3,1).Value2 or '').casefold():
            raise ValueError('Нужна двухуровневая шапка приложения, A3:A4.')
        # Preserve A1 title, A2 object, original widths, merges and styles.
        # All business data are overwritten ONLY in the disposable copy.
        sheet.Range('A'+str(first)+':G'+str(last)).Value2=tuple(synthetic)
        sheet.Range('A'+str(first)+':G'+str(last)).Borders.LineStyle=1
        sheet.PageSetup.PrintArea='$A$1:$G$'+str(last)
        formatted_rows=fit_appendix_body(sheet,first,last)
        set_appendix_print_titles(sheet,first)
        # Numbers 11/22/33 correspond to actual long-name rows 15/26/37.
        height_samples={str(n):float(sheet.Rows(n).RowHeight)
                        for n in (5,10,15,26,37,66,114)}
        if formatted_rows!=len(synthetic) or max(height_samples.values())<=18:
            raise RuntimeError('Длинные наименования не увеличили высоту строк.')
        book.Save()
    finally:
        book.Close(SaveChanges=False)

    proof=app.Workbooks.Open(
        Filename=str(scratch),ReadOnly=True,UpdateLinks=0,AddToMru=False,
    )
    try:
        sheet=proof.Worksheets(1)
        verify_appendix_print_layout(sheet,'3:4')
        for row in (3,4):
            if float(sheet.Rows(row).RowHeight)<28:
                raise RuntimeError('После Excel Save/Open заголовок снова низкий.')
        proof.ExportAsFixedFormat(
            Type=0,Filename=str(result),Quality=0,
            IncludeDocProperties=True,IgnorePrintAreas=False,
            OpenAfterPublish=False,
        )
    finally:
        proof.Close(SaveChanges=False)

    pages=audit_pdf_pages(result,synthetic[0][0],synthetic[-1][0])
    if sha(source)!=original:
        raise RuntimeError('Изменён исходный Excel: '+str(source))

    return {
        'source':str(source),
        'original_sha256':original,
        'pdf_pages':len(pages),
        'page_audit':pages,
        'row_height_samples_pt':height_samples,
        'excel_print_title_rows':'3:4',
        'diagnostic_pdf':str(result),
        'original_unchanged':True,
        'status':'PENDING_VISUAL_REVIEW',
    }


def main(first,akz,output):
    first=Path(first).resolve(strict=True)
    akz=Path(akz).resolve(strict=True)
    if first==akz or any(p.suffix.lower()!='.xlsx' for p in (first,akz)):
        raise ValueError('Нужны два разных проверенных XLSX.')
    target=Path(output).resolve() if output else None
    if target is not None:
        target.mkdir(parents=True,exist_ok=False)
    with tempfile.TemporaryDirectory(prefix='vzmk-pr14-print-') as temp:
        folder=target if target else Path(temp)
        with excel_session() as app:
            reports=[
                export_copy(app,first,folder,'appendix-one'),
                export_copy(app,akz,folder,'akz-appendix'),
            ]
        audit={
            'status':'PENDING_VISUAL_REVIEW',
            'technical_checks_passed':True,
            'human_review_required':True,
            'visual_review_rule':'Inspect EVERY PDF page, especially page 3: '
                                 'all heading cells, no overlap/cropping; '
                                 'read all seven columns and long row names.',
            'reports':reports,
        }
        if target is not None:
            (folder/'PRINT_REVIEW.json').write_text(
                json.dumps(audit,ensure_ascii=False,indent=2),encoding='utf-8')
        return audit


if __name__=='__main__':
    parser=argparse.ArgumentParser()
    parser.add_argument('--appendix-one',required=True)
    parser.add_argument('--akz-appendix',required=True)
    parser.add_argument('--output',help='NEW directory for PDF screenshots and JSON')
    args=parser.parse_args()
    try:
        print(json.dumps(main(args.appendix_one,args.akz_appendix,args.output),
                         ensure_ascii=False,indent=2))
    except BaseException as exc:
        print('FAIL '+type(exc).__name__+': '+str(exc),file=sys.stderr)
        raise
