"""Copy the user's FIRST QUALITY DOCX, changing ONLY yellow-marked fields.

Exact first Garage Kazakhstan document has six marked regions:
passport number, shipped mass, manufacturing end, act references in points
13 and 14, and signature date. All other text/style/paragraphs/tables
stay as provided, including the constant first-manufacture start date.

Numbers and all dates come exclusively from approved VZMK registry facts.
The copy remains UNVERIFIED for actual VIK/UZK/AKZ findings until current
results have been explicitly confirmed; never infer evidence from old text.

No source overwrite: copy only to a brand-new destination with SHA proof.
"""
from datetime import date,datetime
from decimal import Decimal
from pathlib import Path
import re

from docx import Document
from docx.enum.text import WD_COLOR_INDEX

from .storage import Stop,sha
from .highlighted_reference import _save_preserving_parts

NUMBER=re.compile(r'^[1-9][0-9]*(?:/[1-9][0-9]*)?/20[0-9]{2}$')
EXACT_DATE=re.compile(r'^[0-3][0-9]\.[01][0-9]\.20[0-9]{2}$')
DATE_RUN=re.compile(r'^[0-3][0-9]\.[01][0-9]\.20[0-9]{2}$')


def _date(raw,label):
    if isinstance(raw,datetime):raw=raw.date()
    if isinstance(raw,date):return raw.strftime('%d.%m.%Y')
    try: day=date.fromisoformat(str(raw))
    except (ValueError,TypeError) as exc:
        raise Stop('Документ о качестве: дата '+label+
                   ' отсутствует в подтверждённом реестре.') from exc
    return day.strftime('%d.%m.%Y')


def _one_highlighted_cell(table,row,column,pattern):
    if row>=len(table.rows) or column>=len(table.rows[row].cells):
        raise Stop('Документ о качестве: изменена структура первого бланка.')
    cell=table.rows[row].cells[column]
    matches=[
        run for paragraph in cell.paragraphs for run in paragraph.runs
        if run.font.highlight_color==WD_COLOR_INDEX.YELLOW
        and re.fullmatch(pattern,run.text.strip())
    ]
    if not matches and not any(run.font.highlight_color is not None for paragraph in cell.paragraphs for run in paragraph.runs):
        # A later issued sample has the same agreed cells without editing
        # highlights. Exact coordinates, unique runs and whole shape still gate it.
        matches=[run for paragraph in cell.paragraphs for run in paragraph.runs if re.fullmatch(pattern,run.text.strip())]
    if len(matches)!=1:
        raise Stop('Документ о качестве: не найдено ровно одно '
                   'жёлтое поле в строке '+str(row)+': '+str(column))
    return matches[0]


ACT_NUMBER=re.compile(r'^[1-9]\d*/20\d\d(?:-[1-3])?$')
EXTRA_ACT=re.compile(r'^; Сертификат качества антикоррозийной обработки № ([1-9]\d*/20\d\d(?:-[1-3])?) от \d\d\.\d\d\.20\d\d г\.$')


def is_exact_first_quality(first):
    """Read-only shape recognition. Never adapt arbitrary historical DOCX."""
    try:
        doc=Document(first)
        if len(doc.tables)<2:return False
        t=doc.tables[0]
        if len(t.rows)<54:return False
        for row,col,pattern in (
            (1,9,r'[1-9]\d*/20\d\d'),
            (16,6,r'[\d\s]+,\d\d'),
            (18,8,r'\d\d\.\d\d\.20\d\d'),
        ):
            _one_highlighted_cell(t,row,col,pattern)
        _one_highlighted_cell(doc.tables[1],4,0,
                              r'\d\d\.\d\d\.20\d\d')
        for row in (50,53):
            cell=t.rows[row].cells[0]
            paragraph=cell.paragraphs[0]
            value=paragraph.text
            if (value.count('Акт ВИК №')!=2
                    or value.count('Акт УЗК №')!=1
                    or not 1<=value.count('Сертификат качества антикоррозийной обработки №')<=4):
                return False
            yellow=[run for run in paragraph.runs
                    if run.font.highlight_color==WD_COLOR_INDEX.YELLOW]
            if not yellow:yellow=paragraph.runs
            numbers=[run for run in yellow if ACT_NUMBER.fullmatch(run.text)]
            dates=[run for run in yellow if DATE_RUN.fullmatch(run.text)]
            if len(numbers)!=4 or len(dates)!=4:return False
            extras=[r for r in paragraph.runs if EXTRA_ACT.fullmatch(r.text)]
            if value.count('Сертификат качества антикоррозийной обработки №')!=len(extras)+1:return False
        return True
    except Exception:
        return False


def _fill_highlighted(run,value):
    run.text=value
    run.font.highlight_color=None


def _appendix_acts(paragraph,original_number,number,vik,uzk,akz,*,variants=None,uzk_required=True):
    runs=[r for r in paragraph.runs
          if r.font.highlight_color==WD_COLOR_INDEX.YELLOW]
    if not runs:runs=list(paragraph.runs)
    for extra in list(paragraph.runs):
        match=EXTRA_ACT.fullmatch(extra.text)
        if match:
            if match[1].split('-')[0]!=original_number:raise Stop('Ссылка АКЗ относится к другому исходному паспорту.')
            paragraph._p.remove(extra._r)
    numbers=[r for r in runs if ACT_NUMBER.fullmatch(r.text)]
    if any(r.text.split('-')[0]!=original_number for r in numbers):raise Stop('Ссылки актов относятся к другому исходному паспорту.')
    dates=[r for r in runs if DATE_RUN.fullmatch(r.text)]
    if len(numbers)!=4 or len(dates)!=4:
        raise Stop('Документ о качестве: нарушены четыре '
                   'выделенных номера или даты актов.')
    for run in numbers:run.text=number
    for run,value in zip(dates,(vik,vik,uzk,akz)):run.text=value
    from .akz_from_first import act_number
    from copy import deepcopy
    variants=[0] if variants is None else sorted(set(variants))
    if not variants:raise Stop('Документ о качестве: нет подтверждённого вида АКЗ для ссылки.')
    numbers[-1].text=act_number(number,variants[0])
    for variant in variants[1:]:
        # Append only within the agreed variable references paragraph;
        # retain all original paragraph, run styles and surrounding cells.
        run=deepcopy(numbers[-1]._r)
        for text in run.findall('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t'):
            text.text=('; Сертификат качества антикоррозийной обработки № '+
                       act_number(number,variant)+' от '+akz+' г.')
        for mark in list(run.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}highlight')):mark.getparent().remove(mark)
        paragraph._p.append(run)
    # The entire original act reference line was marked, not only numbers.
    # Clear editing highlights in NEW COPY while retaining fonts/spacing.
    for run in runs:run.font.highlight_color=None
    if not uzk_required:
        # Remove only the current UZK reference, across its original runs.
        # Retain the VIK/AKZ text, fonts, paragraph and table geometry.
        pattern=r'Акт УЗК № '+re.escape(number)+r' от '+re.escape(uzk)+r' г\., '
        matches=list(re.finditer(pattern,paragraph.text))
        if len(matches)!=1:raise Stop('Документ о качестве: неоднозначная ссылка УЗК для исключения.')
        start,end=matches[0].span();offset=0
        for run in paragraph.runs:
            text=run.text
            low=max(0,start-offset);high=min(len(text),end-offset)
            if low<high:run.text=text[:low]+text[high:]
            offset+=len(text)


def _manufacturing_start(table,value):
    if value is None:return
    printed=_date(value,'начала изготовления первого паспорта')
    seen=set();matches=[]
    for cell in table.rows[17].cells:
        if id(cell._tc) in seen:continue
        seen.add(id(cell._tc))
        for paragraph in cell.paragraphs:
            for run in paragraph.runs:
                if re.fullmatch(r'\d\d\.\d\d\.20\d\d',run.text.strip()):matches.append(run)
    if len(matches)!=1:raise Stop('Документ о качестве: неоднозначное поле начала изготовления.')
    matches[0].text=printed


def create_quality_copy(source,target,*,number,end_date,vik_date,
                        uzk_date,akz_date,mass_kg,start_date=None,variants=None,uzk_required=True):
    source=Path(source).resolve(strict=True)
    target=Path(target).resolve(strict=False)
    if (source.suffix.lower()!='.docx' or target.suffix.lower()!='.docx'
            or source==target or target.exists() or not target.parent.is_dir()):
        raise Stop('Документ о качестве: требуется отдельный новый '
                   'DOCX, существующий файл не перезаписываем.')
    if not isinstance(number,str) or not NUMBER.fullmatch(number):
        raise Stop('Документ о качестве: номер только из ВЗМК.xlsx.')
    result_dates={
        key:_date(raw,key) for key,raw in (
            ('окончания изготовления',end_date),
            ('ВИК',vik_date),('УЗК',uzk_date),('АКЗ',akz_date),
        )
    }
    try:mass=Decimal(str(mass_kg))
    except Exception as exc:
        raise Stop('Документ о качестве: некорректная масса ВОМ.') from exc
    if not mass.is_finite() or mass<=0:
        raise Stop('Документ о качестве: отсутствует положительная масса.')
    mass_text=f'{mass.quantize(Decimal("0.01")):,.2f}'.replace(',',' ').replace('.',',')
    initial=sha(source)
    if not is_exact_first_quality(source):
        raise Stop('Документ о качестве: первый бланк не совпадает '
                   'с утверждённым отмеченным образцом.')
    doc=Document(source)
    t=doc.tables[0]
    num=_one_highlighted_cell(t,1,9,r'[1-9]\d*/20\d\d')
    source_number=num.text
    _fill_highlighted(num,number)
    _fill_highlighted(
        _one_highlighted_cell(t,16,6,r'[\d\s]+,\d\d'),mass_text)
    _fill_highlighted(
        _one_highlighted_cell(t,18,8,r'\d\d\.\d\d\.20\d\d'),
        result_dates['окончания изготовления'])
    for row in (50,53):
        _appendix_acts(t.rows[row].cells[0].paragraphs[0],
                       source_number,number,
                       result_dates['ВИК'],result_dates['УЗК'],
                       result_dates['АКЗ'],variants=variants,uzk_required=uzk_required)
    _manufacturing_start(t,start_date)
    _fill_highlighted(
        _one_highlighted_cell(doc.tables[1],4,0,
                              r'\d\d\.\d\d\.20\d\d'),
        result_dates['окончания изготовления'])
    _save_preserving_parts(doc,source,target)
    if sha(source)!=initial:
        raise Stop('Документ о качестве: изменён исходный первый паспорт.')
    return {'file':str(target),'number':number,'mass_kg':str(mass),
            'dates':result_dates,
            'source_sha256':initial,
            'status':'PREVIEW_FROM_HISTORICAL_FORM_UNVERIFIED_CURRENT_CONTROL'}
