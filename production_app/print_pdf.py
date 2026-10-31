"""One native-Office PDF for printing the completed passport package.

Order: quality, Appendix No. 1, VIK, UZK, then act+appendix of each
applicable AKZ variant (0..3). Native Word/Excel export retains the original
print layout; temporary individual PDFs are never published.

The PDF is created from FINAL release copies before registry commit.
Any conversion/validation failure prevents the release from being accepted.
"""
import os
import re
import shutil
import tempfile
from io import BytesIO
from pathlib import Path

from pypdf import PdfReader, PdfWriter

from passport_app.excel import excel_session
from passport_app.vik import word_session, open_word
from .storage import Stop, pins, verify_pins, sha
from .office import PREVIEW_MARK, word_preview_mark

BASE_ROLES = ('quality', 'appendix', 'vik', 'uzk')
ROLE_EXTENSIONS = {
    'quality': '.docx', 'appendix': '.xlsx',
    'vik': '.docx', 'uzk': '.docx',
}
AKZ_ROLE = re.compile(r'akz(?:_appendix)?[0-3]\Z')


def ordered_sources(outputs, *, partial_draft=False, uzk_required=True):
    """Return complete ordered components; refuse duplicates and orphaned AKZ."""
    if not isinstance(outputs, list):
        raise Stop('Неверный состав документов для общего PDF.')
    roles = {}
    for item in outputs:
        if not isinstance(item, dict) or not isinstance(item.get('role'), str):
            raise Stop('Документ без роли в комплекте PDF.')
        role = item['role']
        if role in roles:
            raise Stop('Повтор роли документа при сборке PDF: ' + role)
        if role not in BASE_ROLES and not AKZ_ROLE.fullmatch(role):
            raise Stop('Неизвестный документ при сборке PDF: ' + role)
        path = Path(item.get('path', ''))
        expected = ROLE_EXTENSIONS.get(role, '.xlsx')
        if path.suffix.lower() != expected or not path.is_file():
            raise Stop('Нет готового ' + expected + ' для роли ' + role + ': ' + str(path))
        roles[role] = path

    if not uzk_required and 'uzk' in roles:
        raise Stop('При нулевом плане УЗК акт УЗК не входит в комплект PDF.')
    missing = [role for role in BASE_ROLES if role not in roles and (role!='uzk' or uzk_required)]
    if missing and not partial_draft:
        raise Stop('В комплекте PDF отсутствуют обязательные документы: ' + ', '.join(missing))
    ordered = [dict(role=role, path=roles[role]) for role in BASE_ROLES if role in roles]
    if not ordered:
        raise Stop('Нет доступных документов для общего PDF.')
    for variant in range(4):
        act, appendix = f'akz{variant}', f'akz_appendix{variant}'
        if (act in roles) != (appendix in roles):
            raise Stop('Для АКЗ №' + str(variant) + ' требуются акт и его приложение.')
        if act in roles:
            ordered.extend((dict(role=act, path=roles[act]),
                            dict(role=appendix, path=roles[appendix])))
    return ordered


def _export_word(items):
    if not items:
        return
    with word_session() as word:
        for item in items:
            with tempfile.TemporaryDirectory(prefix='.word-print-', dir=item['pdf'].parent) as tmp:
                source = item['source']
                if item.get('preview'):
                    # Use the SAME XML-only paragraph removal as release_copies.
                    # Word Range.Delete can retain paragraph spacing and move
                    # the first table. Only this private export copy is changed.
                    source = Path(tmp) / source.name
                    shutil.copyfile(item['source'], source)
                    word_preview_mark(source, remove=True)
                doc = open_word(word, source, readonly=True)
                try:
                    if item['role']=='uzk':_fit_uzk_executor(doc)
                    doc.ExportAsFixedFormat(
                        OutputFileName=str(item['pdf']), ExportFormat=17,
                        OpenAfterExport=False, OptimizeFor=0,
                    )
                finally:
                    doc.Close(SaveChanges=False)


def _fit_uzk_executor(doc):
    """Native pagination check for the recognized historical UZK only.

    Word ignores the keep chain at this merged table boundary for some
    row counts. Force a break BEFORE the last mark when its complete signer
    would be orphaned. This is only the unsaved read-only PDF export buffer.
    """
    if doc.Tables.Count<2:return
    table=doc.Tables(1)
    if table.Rows.Count<6:return
    if 'Характеристика сварного соединения' not in table.Rows(5).Range.Text:return
    executor=doc.Tables(2)
    if 'Контроль выполнил' not in executor.Range.Text:return
    last=table.Rows(table.Rows.Count)
    def page_at(position):
        return doc.Range(Start=position,End=position).Information(3)
    def together():
        return (page_at(last.Range.Start)==page_at(last.Range.End-2)
                ==page_at(executor.Range.Start)==page_at(executor.Range.End-2))
    doc.Repaginate()
    if not together():
        last.Range.Paragraphs(1).Format.PageBreakBefore=True
        doc.Repaginate()
        if not together():
            raise Stop('УЗК: последняя марка и полный блок исполнителя '
                       'не помещаются на одном листе. Требуется проверка формы.')


def _export_excel(items):
    if not items:
        return
    with excel_session() as excel:
        for item in items:
            book = excel.Workbooks.Open(
                Filename=str(item['source'].resolve()),
                UpdateLinks=0, ReadOnly=True,
                IgnoreReadOnlyRecommended=True, Notify=False, AddToMru=False,
            )
            try:
                if not book.ReadOnly:
                    raise Stop('Книга для печати PDF не открылась только для чтения: '
                               + str(item['source']))
                if item.get('preview'):
                    for sheet in book.Worksheets:
                        if sheet.PageSetup.CenterHeader!=PREVIEW_MARK:
                            raise Stop('Изменена маркировка предварительного просмотра Excel.')
                        sheet.PageSetup.CenterHeader=''
                # Export the original print areas and page breaks. Never
                # change sheet layout/formatting or save the source workbook.
                book.ExportAsFixedFormat(
                    Type=0, Filename=str(item['pdf']), Quality=0,
                    IncludeDocProperties=True, IgnorePrintAreas=False,
                    OpenAfterPublish=False,
                )
            finally:
                book.Close(SaveChanges=False)


def _pdf_pages(path):
    if not path.is_file() or path.stat().st_size < 20:
        raise Stop('Microsoft Office не создал PDF: ' + str(path))
    with path.open('rb') as handle:
        if handle.read(5) != b'%PDF-':
            raise Stop('Файл не является PDF: ' + str(path))
    try:
        pdf = PdfReader(str(path), strict=False)
        if pdf.is_encrypted:
            raise Stop('PDF требует пароль: ' + str(path))
        pages = len(pdf.pages)
        if pages < 1:
            raise Stop('PDF не содержит страниц: ' + str(path))
        for page in pdf.pages:
            rect = page.mediabox
            if float(rect.width) <= 0 or float(rect.height) <= 0:
                raise Stop('Страница PDF имеет недопустимый размер: ' + str(path))
        return pages
    except Stop:
        raise
    except Exception as exc:
        raise Stop('Ошибка проверки PDF ' + str(path) + ': ' + str(exc)) from exc


def verify_print_pdf(path, *, expected_pages=None):
    """Read-only validation for release and final Windows Office readback."""
    pages = _pdf_pages(Path(path))
    if expected_pages is not None and pages != expected_pages:
        raise Stop('После выпуска изменилось количество страниц сводного PDF.')
    return pages



def verify_package_print(released, package_folder_name, *, uzk_required=True):
    """Validate the print manifest, order, pages and hash before registry commit."""
    if not isinstance(released, dict):
        raise Stop('Отсутствует проверенный манифест выпускных документов.')
    outputs = released.get('outputs')
    if not isinstance(outputs, list):
        raise Stop('Отсутствует список готовых документов.')
    pdf_items = [item for item in outputs if item.get('role') == 'print_pdf']
    if len(pdf_items) != 1:
        raise Stop('В выпуске должен быть ровно один сводный PDF.')
    ordered = ordered_sources([item for item in outputs if item.get('role') != 'print_pdf'],uzk_required=uzk_required)
    path = Path(pdf_items[0]['path'])
    manifest = released.get('print_pdf')
    name = package_folder_name + '.pdf'
    if (not isinstance(manifest, dict) or path.name != name or
            manifest.get('filename') != name):
        raise Stop('Название сводного PDF не совпадает с папкой паспорта.')
    if any(item['path'].parent.resolve() != path.parent.resolve() for item in ordered):
        raise Stop('Документы для печати и сводный PDF должны находиться в одной папке.')

    sections = manifest.get('sections')
    if not isinstance(sections, list) or len(sections) != len(ordered):
        raise Stop('Неполный список разделов сводного PDF.')
    expected_start = 1
    for item, section in zip(ordered, sections):
        if not isinstance(section, dict) or section.get('role') != item['role']:
            raise Stop('Неверный порядок документов в сводном PDF.')
        count = section.get('pages')
        if (type(count) is not int or count < 1 or
                section.get('start_page') != expected_start):
            raise Stop('Неверная нумерация/количество страниц сводного PDF.')
        expected_start += count
    expected_count = expected_start - 1
    if manifest.get('pages') != expected_count:
        raise Stop('Не совпало количество страниц разделов PDF.')
    if sha(path) != manifest.get('sha256'):
        raise Stop('Сводный PDF изменился после формирования.')
    verify_print_pdf(path, expected_pages=expected_count)
    return manifest


def _preview_overlay(page):
    """A Unicode, embedded-font mark in the PDF margin, after pagination."""
    from reportlab.pdfgen import canvas
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    name='VZMKPreviewArial'
    if name not in pdfmetrics.getRegisteredFontNames():
        font=Path(os.environ.get('WINDIR','C:/Windows'))/'Fonts'/'arial.ttf'
        if not font.is_file():raise Stop('Нет шрифта Arial для маркировки PDF.')
        pdfmetrics.registerFont(TTFont(name,str(font)))
    page.transfer_rotation_to_content()
    box=page.cropbox
    width,height=float(box.width),float(box.height)
    stream=BytesIO();c=canvas.Canvas(stream,pagesize=(width,height))
    size=min(9,(width-24)/pdfmetrics.stringWidth(PREVIEW_MARK,name,1))
    c.setFont(name,size);c.setFillColorRGB(.65,0,0)
    # The unshifted historical continuation starts about 15 pt from the top.
    # Keep the label above that text rather than occupying its first line.
    c.drawCentredString(width/2,height-8,PREVIEW_MARK)
    c.save();stream.seek(0)
    overlay=PdfReader(stream).pages[0]
    overlay.add_transformation((1,0,0,1,float(box.left),float(box.bottom)))
    overlay.mediabox=page.mediabox
    overlay.cropbox=page.cropbox
    page.merge_page(overlay)


def create_print_pdf(outputs, release_folder, package_folder_name, *, partial_draft=False, preview=False, uzk_required=True):
    """Return PDF output record + section manifest; no other files published.

    release_folder is a private, newly created directory. package_folder_name
    is already generated by the approved passport output naming function.
    """
    release_folder = Path(release_folder)
    if (not release_folder.is_dir() or
            Path(package_folder_name).name != package_folder_name or
            package_folder_name in ('', '.', '..') or
            any(char in package_folder_name for char in '<>:"/\\|?*')):
        raise Stop('Небезопасное имя папки/место сборки сводного PDF.')
    ordered = ordered_sources(outputs, partial_draft=partial_draft, uzk_required=uzk_required)
    if any(item['path'].parent.resolve() != release_folder.resolve()
           for item in ordered):
        raise Stop('Общий PDF разрешено создавать только из выпускных копий в одной папке.')
    target = release_folder / (package_folder_name + '.pdf')
    if target.exists():
        raise Stop('Сводный PDF уже существует; перезапись запрещена: ' + str(target))

    originals = pins([item['path'] for item in ordered])
    sections = []
    # All per-document PDFs stay in a private temporary directory, removed
    # after merge. PDF is not visible in the release folder until fully valid.
    with tempfile.TemporaryDirectory(prefix='.pdf-export-', dir=release_folder) as tmp:
        stage = Path(tmp)
        conversions = []
        for pos, entry in enumerate(ordered):
            conversions.append({
                'role': entry['role'],
                'source': entry['path'],
                'pdf': stage / (f'{pos:02d}-{entry["role"]}.pdf'),
                'preview': preview,
            })
        _export_word([item for item in conversions if item['source'].suffix.lower() == '.docx'])
        _export_excel([item for item in conversions if item['source'].suffix.lower() == '.xlsx'])
        verify_pins(originals)

        writer = PdfWriter()
        page_start = 1
        try:
            for item in conversions:
                count = _pdf_pages(item['pdf'])
                writer.append(str(item['pdf']), import_outline=False)
                sections.append({'role': item['role'], 'pages': count, 'start_page': page_start})
                page_start += count
            if preview:
                for page in writer.pages:_preview_overlay(page)
            # Exclusive creation: never silently overwrite an existing file.
            with target.open('xb') as out:
                try:
                    writer.write(out)
                    out.flush()
                    os.fsync(out.fileno())
                except BaseException:
                    out.close()
                    target.unlink(missing_ok=True)
                    raise
        finally:
            writer.close()

    verify_pins(originals)
    pages = verify_print_pdf(target, expected_pages=sum(s['pages'] for s in sections))
    return {
        'output': {'role': 'print_pdf', 'path': str(target)},
        'manifest': {
            'filename': target.name,
            'sha256': sha(target),
            'pages': pages,
            'sections': sections,
        },
    }
