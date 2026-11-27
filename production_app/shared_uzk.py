"""One shared VZMK ultrasonic inspection XLSM for all projects.

Reading/planning/drafts never write. Official accounting creates a separate
validated candidate; the existing transaction coordinator alone publishes it.

The shared workbook contains *remaining quantities*, NOT inspected weld
outcomes. It can populate the planned UZK quantity for current VOM marks,
but cannot assert GОДЕН, dates, actual weld length or personnel.

Contract from the approved Stage 5 project:
 sheet 'УЗК', first-row columns 'Заказ', 'Марка элемента', 'Кол-во (шт)'.
 Blank numeric cell -> 1; exact identical duplicate doesn't increase stock;
 conflicting duplicate blocks. Selected project's rows only.
"""
from decimal import Decimal, InvalidOperation
from collections import Counter
from bisect import bisect_left
from pathlib import Path
from openpyxl import load_workbook
from .storage import Stop, sha, safe, post_commit_path


HEADERS=('Заказ','Марка элемента','Кол-во (шт)')


def _key(raw):
    return ' '.join(str(raw or '').replace('ё','е').casefold().split())


def _qty(raw,row,*,allow_zero=False):
    if raw is None or isinstance(raw,str) and not raw.strip():
        return 1
    if isinstance(raw,bool) or isinstance(raw,str) and raw.startswith('='):
        raise Stop('Общий УЗК: неверное количество в строке '+str(row))
    try:
        q=Decimal(str(raw).strip().replace(',','.'))
    except (ValueError,InvalidOperation):
        raise Stop('Общий УЗК: нечисловое количество в строке '+str(row))
    from .shipment_input import EXCEL_MAX
    if (not q.is_finite() or q<0 or not allow_zero and q==0 or q>EXCEL_MAX
            or q!=q.to_integral_value() or Decimal(format(q,'.15g'))!=q):
        raise Stop('Общий УЗК: количество должно быть положительным целым, '
                   'строка '+str(row))
    return int(q)


def find_shared_uzk(cfg,*,require=True):
    """A single explicitly selected file, or the unambiguous standard sibling.

    Its location does not depend on the currently selected project, and
    nothing is saved or written during discovery.
    """
    explicit=str(cfg.get('uzk_registry') or '').strip()
    if not explicit:
        registry=cfg.get('registry')
        if not registry:
            if require:raise Stop('Выберите реестр ВЗМК перед поиском общего УЗК.')
            return None
        candidate=Path(registry).parent/'УЗК.xlsm'
    else:
        candidate=Path(explicit)
    candidate=safe(candidate)
    if not candidate.is_file():
        if require:
            raise Stop('Общий реестр УЗК.xlsm не найден: '+str(candidate)+
                       '. Выберите его один раз в «Настройки путей».')
        return None
    if candidate.suffix.lower()!='.xlsm':
        raise Stop('Общий реестр УЗК должен быть файлом *.xlsm.')
    return candidate


def read_shared_uzk(path,*,project=None,marks=None,post_commit=False):
    """Read only values; reject formula quantities, duplicate contradictory rows."""
    path=post_commit_path(path) if post_commit else safe(path)
    before=sha(path)
    book=load_workbook(path,read_only=True,data_only=False,keep_links=False)
    rows={}
    project_key=_key(project) if project is not None else None
    selected_marks={_key(mark) for mark in marks} if marks is not None else None
    try:
        if 'УЗК' not in book.sheetnames:
            raise Stop('В общем XLSM отсутствует лист «УЗК».')
        sheet=book['УЗК']
        header=next(sheet.iter_rows(min_row=1,max_row=1,max_col=3,
                                    values_only=True),())
        if tuple(_key(h) for h in header)!=tuple(_key(h) for h in HEADERS):
            raise Stop('Неверные заголовки УЗК: ожидаются Заказ | Марка '
                       'элемента | Кол-во (шт).')
        for no,values in enumerate(
            sheet.iter_rows(min_row=2,max_col=3,values_only=True),start=2
        ):
            order,mark,qty=values
            if project_key is not None:
                if isinstance(order,str) and order.strip() and not order.startswith('='):
                    if _key(order)!=project_key:continue
                elif selected_marks is not None and _key(mark) not in selected_marks:
                    continue  # Unassigned rows of other marks remain opaque and untouched.
            if all(x is None or isinstance(x,str) and not x.strip()
                   for x in values):
                continue
            if (not isinstance(order,str) or not order.strip() or
                    not isinstance(mark,str) or not mark.strip() or
                    order.startswith('=') or mark.startswith('=')):
                raise Stop('Общий УЗК: заказ/марка отсутствует или неверен '
                           'в строке '+str(no))
            count=_qty(qty,no)
            key=(_key(order),_key(mark))
            if key in rows:
                if rows[key]['remaining']!=count:
                    raise Stop('Общий УЗК: противоречивые остатки у марки '+
                               mark+', строки '+str(rows[key]['rows'][0])+
                               ' и '+str(no))
                rows[key]['rows'].append(no)
            else:
                rows[key]={'project':order,'mark':mark,'remaining':count,
                           'rows':[no]}
    finally:
        book.close()
    if sha(path)!=before:
        raise Stop('Общий реестр УЗК изменился во время чтения.')
    return rows


def planned_control(path,project,vom):
    """Compute min(shipped, remaining) independently for every VOM mark.

    The first encounter determines display name; repeated same mark in VOM
    accumulates quantity. Never sum different marks or duplicated registry
    rows. Missing registry mark => no planned UZK, NOT a negative verdict.
    """
    shipped={}
    labels={}
    for row in vom.rows:
        mark=str(row[0])
        key=_key(mark)
        if not key:
            raise Stop('Марка ВОМ не определена.')
        quantity=_qty(row[2], 'ВОМ '+mark)
        shipped[key]=shipped.get(key,0)+quantity
        labels.setdefault(key,mark)
    rows=read_shared_uzk(path,project=project,marks=labels.values())
    output={}
    for key,quantity in shipped.items():
        rec=rows.get((_key(project),key))
        required=min(quantity,rec['remaining']) if rec else 0
        output[labels[key]]={
            'mark':labels[key],
            'vom_quantity':quantity,
            'registry_remaining':rec['remaining'] if rec else None,
            'uzk_quantity':required,
            'matched':bool(rec),
            'manual_limit':quantity if not rec else 0,
            'registry_rows':list(rec['rows']) if rec else [],
        }
    return output


def manual_control_input(raw,mark,control):
    """The visible total includes retained controls and the protected registry plan."""
    base=control['uzk_quantity']
    observed=base if raw is None or not str(raw).strip() else _qty(raw,'ВОМ '+mark,allow_zero=True)
    extra=observed-base
    if not 0<=extra<=control.get('manual_limit',0):
        raise Stop(f'УЗК: марка {mark}, допустимое количество от {base} до '
                   f'{base+control.get("manual_limit",0)} шт.')
    return extra


def current_design_quantity(row,control):
    auto=control.get('registry_added',control['uzk_quantity'])
    return auto+int(row.get('uzk_manual_requested',0))


def verify_planned_quantities(facts,plan,*,accept_manual=False):
    """Cannot change a computed quantity to skip selected current inspection."""
    expected={name:int(x['uzk_quantity']) for name,x in plan.items()}
    marks=[str(r.get('mark','')) for r in facts.get('marks',[])]
    missing=[m for m in expected if m not in marks]
    extra=list(dict.fromkeys(m for m in marks if m not in expected))
    frequencies=Counter(marks)
    duplicates=list(dict.fromkeys(m for m in marks if frequencies[m]>1))
    if missing or extra or duplicates:
        raise Stop('УЗК: состав марок не совпадает с текущей ВОМ; ожидалось '+str(len(expected))+
                   ', найдено '+str(len(marks))+'.\nОтсутствуют: '+(', '.join(missing) or 'нет')+
                   '.\nЛишние: '+(', '.join(extra) or 'нет')+'.\nПовторы: '+(', '.join(duplicates) or 'нет')+'.')
    for row in facts.get('marks',[]):
        mark=str(row.get('mark',''))
        if mark not in expected:
            raise Stop('УЗК: неизвестная марка текущей ВОМ '+mark)
        raw=row.get('uzk_quantity')
        if raw is None or isinstance(raw,bool) or isinstance(raw,str) and not raw.strip():
            raise Stop('УЗК: не заполнено действующее количество для марки '+mark)
        observed=_qty(raw,'ВОМ '+mark,allow_zero=True)
        choice=plan[mark]
        manual=row.get('uzk_manual_requested',0) if accept_manual else 0
        extra=0
        if manual not in (0,'0'):
            extra=manual_control_input(raw,mark,choice)
            if manual is None or not str(manual).strip() or not extra or extra!=_qty(manual,mark):
                raise Stop('УЗК: ручное количество не соответствует текущей ВОМ: '+mark)
        if observed!=expected[mark]+extra:
            raise Stop('УЗК: количество по марке '+mark+' должно быть '+
                       str(expected[mark])+' согласно общему реестру. '
                       'Изменять его вручную для выпуска нельзя.')
        if accept_manual:
            if choice.get('replacement_min',0)>choice.get('registry_added',0)+extra:
                raise Stop('УЗК: подтвердите контроль заменяемых изделий марки '+mark+'.')
            previous_manual=choice.get('manual_retained',0)
            row['uzk_quantity']=observed
            row['uzk_manual_requested']=extra
            row['uzk_registry_quantity']=expected[mark]-previous_manual
            row['uzk_manual_quantity']=previous_manual+extra
    return True

def _debit_rows(rows,project,quantities):
    updates={}
    deletions=[]
    for mark,value in quantities.items():
        if value in (0,'0'):continue
        count=_qty(value,mark);rec=rows.get((_key(project),_key(mark)))
        if not rec:raise Stop('УЗК: отсутствует списываемая марка '+mark)
        if len(rec['rows'])!=1:raise Stop('УЗК: дубли строк марки '+mark+'; запись запрещена.')
        if count>rec['remaining']:raise Stop('УЗК: недостаточный остаток '+mark)
        row=rec['rows'][0];left=rec['remaining']-count
        if left:updates[row]=left
        else:deletions.append(row)
    return updates,deletions


def _uzk_xml(archive):
    import posixpath
    from lxml import etree

    x='{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
    r='{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
    book=etree.fromstring(archive.read('xl/workbook.xml'))
    sheet=next(node for node in book.find(x+'sheets') if node.get('name')=='УЗК')
    rels=etree.fromstring(archive.read('xl/_rels/workbook.xml.rels'))
    ref=next(rel.get('Target') for rel in rels if rel.get('Id')==sheet.get(r+'id'))
    part=ref.lstrip('/') if ref.startswith('/') else posixpath.normpath('xl/'+ref)
    return book,part,etree.fromstring(archive.read(part))


def _debit_deletions_require_excel(archive,book,tree,sheet_path):
    import re
    from lxml import etree

    x='{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
    unsupported=('mergeCells','tableParts','drawing','legacyDrawing','hyperlinks')
    complex_parts={tag for tag in unsupported if tree.find(x+tag) is not None}
    if complex_parts:
        if not complex_parts <= {'drawing','legacyDrawing'}:
            raise Stop('УЗК: удаление строк в сложной структуре требует отдельной проверки; исходник не изменён.')
        return True
    for cf in tree.findall(x+'conditionalFormatting'):
        if any(etree.QName(node).localname=='formula' for node in cf.iter()):
            raise Stop('УЗК: формулы условного форматирования зависят от строк; запись остановлена.')
    validations=tree.find(x+'dataValidations')
    if validations is not None:
        for validation in validations:
            for node in validation:
                if etree.QName(node).localname.startswith('formula') and not str(node.text or '').startswith('"'):
                    raise Stop('УЗК: проверка данных содержит ссылки на строки; запись остановлена.')
    for name in archive.namelist():
        if name.startswith('xl/worksheets/') and name.endswith('.xml'):
            other=tree if name==sheet_path else etree.fromstring(archive.read(name))
            for formula in other.iter(x+'f'):
                if name==sheet_path or 'УЗК' in (formula.text or '').upper():
                    raise Stop('УЗК: найдены формулы со ссылками на удаляемые строки; запись остановлена.')
    for defined in book.iter(x+'definedName'):
        if 'УЗК' in (defined.text or '').upper():
            if defined.get('name') not in ('_xlnm._FilterDatabase','_xlnm.Print_Area') or not re.fullmatch(r"(?:'УЗК'|УЗК)![A-Z$0-9:]+",defined.text or ''):
                raise Stop('УЗК: именованные диапазоны зависят от удаляемых строк; запись остановлена.')
    return False


def decrement_candidate(source,target,project,quantities):
    """Patch ONLY the data sheet XML, preserving VBA and all other ZIP parts.

    Do not round-trip an XLSM through openpyxl (which can strip extensions).
    Complex references needing cross-sheet edits are refused, never guessed.
    """
    import re
    from zipfile import ZipFile
    from lxml import etree as E

    source=safe(source);target=safe(target)
    if source==target or target.exists():raise Stop('УЗК: нужен отдельный новый кандидат XLSM.')
    original=sha(source);rows=read_shared_uzk(source,project=project,marks=quantities)
    updates,deletions=_debit_rows(rows,project,quantities)
    if not updates and not deletions:
        # A zero plan changes neither quantities nor any opaque XLSM part.
        import shutil
        shutil.copyfile(source,target)
        if sha(source)!=original or sha(target)!=original:raise Stop('УЗК изменился при копировании нулевого плана.')
        return {'sha256':original,'deleted_rows':[],'updated_rows':{},'sheet_xml':None,'changed_parts':[]}
    X='{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
    with ZipFile(source) as z:
        book,sheet_path,tree=_uzk_xml(z)
        data=tree.find(X+'sheetData')
        if deletions and _debit_deletions_require_excel(z,book,tree,sheet_path):
            return _native_drawing_candidate(source,target,original,sheet_path,updates,deletions)
        # Sorting once avoids a full scan of all deleted rows for each cell.
        removed=sorted(deletions)
        removed_set=set(deletions)
        def shifted(no):return no-bisect_left(removed,no)
        for row in list(data):
            no=int(row.get('r'))
            if no in deletions:data.remove(row);continue
            if no in updates:
                cell=next((c for c in row if c.get('r')=='C'+str(no)),None)
                if cell is None:cell=E.SubElement(row,X+'c',r='C'+str(no))
                cell.attrib.pop('t',None)
                for element in list(cell):
                    if element.tag in (X+'f',X+'v',X+'is'):cell.remove(element)
                E.SubElement(cell,X+'v').text=str(updates[no])
            row.set('r',str(shifted(no)))
            for c in row:
                if c.get('r'):c.set('r',re.sub(r'\d+$',str(shifted(no)),c.get('r')))
        def range_shift(value):
            result=[]
            for token in value.split():
                match=re.fullmatch(r'(\$?[A-Z]+)(\$?)([0-9]+)(?::(\$?[A-Z]+)(\$?)([0-9]+))?',token)
                if not match:raise Stop('УЗК: сложный диапазон строки '+token)
                low=int(match[3]);high=int(match[6] or low)
                if low==1 and high==1048576:result.append(token);continue
                # A sparse set of removed rows must not trigger an O(range)
                # walk (Excel ranges can span over one million rows).
                new_low=low
                while new_low<=high and new_low in removed_set:new_low+=1
                if new_low>high:continue
                new_high=high
                while new_high>=new_low and new_high in removed_set:new_high-=1
                first=match[1]+match[2]+str(shifted(new_low))
                last=(match[4] or match[1])+(match[5] or '')+str(shifted(new_high))
                result.append(first+(':'+last if match[6] else ''))
            return ' '.join(result)
        for tag in ('dimension','autoFilter'):
            el=tree.find(X+tag)
            if el is not None and el.get('ref'):el.set('ref',range_shift(el.get('ref')))
        for cf in list(tree.findall(X+'conditionalFormatting')):
            ref=range_shift(cf.get('sqref',''))
            if ref:cf.set('sqref',ref)
            else:tree.remove(cf)
        validations=tree.find(X+'dataValidations')
        if validations is not None:
            for validation in list(validations):
                ref=range_shift(validation.get('sqref',''))
                if ref:validation.set('sqref',ref)
                else:validations.remove(validation)
            validations.set('count',str(len(validations)))
            if not len(validations):tree.remove(validations)
        changed_parts={sheet_path:E.tostring(tree,encoding='UTF-8',xml_declaration=True)}
        for defined in book.iter(X+'definedName'):
            if deletions and 'УЗК' in (defined.text or '').upper():
                prefix,ref=defined.text.split('!',1);changed=prefix+'!'+range_shift(ref)
                if changed!=defined.text:
                    defined.text=changed;changed_parts['xl/workbook.xml']=E.tostring(book,encoding='UTF-8',xml_declaration=True)
        with ZipFile(target,'w') as out:
            for item in z.infolist():out.writestr(item,changed_parts.get(item.filename,z.read(item.filename)))
    if sha(source)!=original:raise Stop('УЗК изменился при подготовке списания.')
    after=read_shared_uzk(target,project=project,marks=quantities)
    for mark,value in quantities.items():
        if value in (0,'0'):continue
        key=(_key(project),_key(mark));left=rows[key]['remaining']-int(value)
        if left and (key not in after or after[key]['remaining']!=left) or not left and key in after:
            raise Stop('УЗК: не подтверждено списание '+mark)
    for key,rec in rows.items():
        if key[0]!=_key(project) or not any(_key(m)==key[1] and q not in (0,'0') for m,q in quantities.items()):
            if key not in after or after[key]['remaining']!=rec['remaining']:raise Stop('УЗК: изменена другая марка.')
    return {'sha256':sha(target),'deleted_rows':deletions,'updated_rows':updates,'sheet_xml':sheet_path,'changed_parts':list(changed_parts)}

def _refund_entries(rows,project,quantities,originals):
    entries=[]
    for mark,qty in quantities.items():
        count=_qty(qty,mark)
        key=(_key(project),_key(mark))
        original=originals.get(mark)
        if (not isinstance(original,dict) or
                (_key(original.get('project')),_key(original.get('mark')))!=key):
            raise Stop('УЗК: не совпала исходная строка марки '+mark+'.')
        current=rows.get(key)
        if current and len(current['rows'])!=1:
            raise Stop('УЗК: дубли строк марки '+mark+'; возврат остановлен.')
        entries.append((key,count,original,current))
    return entries


def _refund_append_extent(tree,mark):
    import re

    x='{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
    dimension=tree.find(x+'dimension')
    extent=re.fullmatch(r'A1:([A-Z]{1,3})([1-9][0-9]*)',dimension.get('ref','')) if dimension is not None else None
    if (extent is None or extent[1] in ('A','B') or
            any(tree.find(x+tag) is not None for tag in ('tableParts','mergeCells'))):
        raise Stop('УЗК: не подтверждено безопасное место для восстановления строки '+mark+'.')
    last=max(int(row.get('r')) for row in tree.find(x+'sheetData'))
    return dimension,extent[1],max(last,int(extent[2]))


def preflight_accounting(source,project,controls,refunds,refund_rows):
    """Check planned changes without writing a candidate or opening Excel."""
    from zipfile import ZipFile

    source=safe(source)
    before=sha(source)
    rows=read_shared_uzk(source,project=project,marks=controls.keys()|refunds.keys())
    updates,deletions=_debit_rows(rows,project,controls)
    removed=set(deletions)
    for key,record in list(rows.items()):
        row=record['rows'][0]
        if row in removed:
            del rows[key]
        elif row in updates:
            rows[key]={**record,'remaining':updates[row]}
    entries=_refund_entries(rows,project,refunds,refund_rows)
    missing=[original['mark'] for _,_,original,current in entries if not current]
    if deletions or missing:
        with ZipFile(source) as archive:
            book,sheet_path,tree=_uzk_xml(archive)
            if deletions:
                _debit_deletions_require_excel(archive,book,tree,sheet_path)
            if missing:
                _,_,last=_refund_append_extent(tree,missing[0])
                new_last=last-len(deletions)+len(missing)
                if new_last>1048576:raise Stop('УЗК: лист заполнен.')
                _extend_plain_refund_filter(book,tree,new_last)
    if sha(source)!=before:
        raise Stop('УЗК изменился во время предварительной проверки.')


def refund_candidate(source,target,project,quantities,originals):
    """Restore confirmed quantities on the private XLSM without shifting rows."""
    from zipfile import ZipFile
    from lxml import etree as E

    source=safe(source);target=safe(target)
    if source==target or target.exists():raise Stop('УЗК: нужен отдельный кандидат возврата.')
    before=sha(source)
    rows=read_shared_uzk(source,project=project,marks=quantities)
    entries=_refund_entries(rows,project,quantities,originals)
    X='{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
    with ZipFile(source) as z:
        book,part,tree=_uzk_xml(z)
        data=tree.find(X+'sheetData')
        by_row={int(row.get('r')):row for row in data}
        expected=dict(rows)
        missing=[original['mark'] for _,_,original,current in entries if not current]
        workbook_changed=False
        if missing:
            dimension,end_col,last=_refund_append_extent(tree,missing[0])
            new_last=last+len(missing)
            if new_last>1048576:raise Stop('УЗК: лист заполнен.')
            workbook_changed=_extend_plain_refund_filter(book,tree,new_last)
            dimension.set('ref','A1:'+end_col+str(new_last))
        else:
            last=max(by_row,default=1)
        for key,count,original,current in entries:
            if current:
                no=current['rows'][0];row=by_row[no]
                count+=current['remaining']
            else:
                last+=1
                no=last
                row=E.SubElement(data,X+'row',r=str(no))
                for col,name in (('A','project'),('B','mark')):
                    cell=E.SubElement(row,X+'c',r=col+str(no),t='inlineStr')
                    val=E.SubElement(E.SubElement(cell,X+'is'),X+'t')
                    val.set('{http://www.w3.org/XML/1998/namespace}space','preserve')
                    val.text=original[name]
            ref='C'+str(no)
            cell=next((c for c in row if c.get('r')==ref),None)
            if cell is None:
                cell=E.Element(X+'c',r=ref)
                # XML cells are stored in ascending column order.
                following=next((c for c in row if c.tag!=X+'c' or
                                len(c.get('r','').rstrip('0123456789'))>1 or c.get('r','')>ref),None)
                row.insert(row.index(following) if following is not None else len(row),cell)
            cell.attrib.pop('t',None)
            for node in list(cell):
                if node.tag in (X+'f',X+'v',X+'is'):cell.remove(node)
            cell.insert(0,E.Element(X+'v'))
            cell[0].text=str(count)
            expected[key]={**(current or original),'rows':[no],'remaining':count}
        changed={part:E.tostring(tree,encoding='UTF-8',xml_declaration=True)}
        if workbook_changed:
            changed['xl/workbook.xml']=E.tostring(book,encoding='UTF-8',xml_declaration=True)
        with ZipFile(target,'x') as out:
            for item in z.infolist():
                payload=changed[item.filename] if item.filename in changed else z.read(item.filename)
                out.writestr(item,payload)
    if sha(source)!=before or read_shared_uzk(target,project=project,marks=quantities)!=expected:
        raise Stop('УЗК: не подтверждён кандидат возврата.')


def _extend_plain_refund_filter(book,tree,new_last):
    """Append under a plain three-column filter and keep its scoped name in sync."""
    import re

    x='{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
    filters=tree.findall(x+'autoFilter')
    properties=tree.find(x+'sheetPr')
    if properties is not None and properties.get('filterMode') not in (None,'0','false'):
        raise Stop('УЗК: для возврата нужна проверка настроек автофильтра.')
    if not filters:
        return False
    auto_filter=filters[0]
    extent=re.fullmatch(r'A1:C([1-9][0-9]*)',auto_filter.get('ref',''))
    allowed_attributes={'ref','{http://schemas.microsoft.com/office/spreadsheetml/2014/revision}uid'}
    if (len(filters)!=1 or extent is None or int(extent[1])>1048576 or
            len(auto_filter) or set(auto_filter.attrib)-allowed_attributes):
        raise Stop('УЗК: для возврата нужна проверка настроек автофильтра.')
    old_last=int(extent[1])
    scope=next(str(i) for i,s in enumerate(book.find(x+'sheets')) if s.get('name')=='УЗК')
    names=book.findall(f'{x}definedNames/{x}definedName')
    local=[n for n in names if n.get('localSheetId')==scope
           and n.get('name')=='_xlnm._FilterDatabase']
    if len(local)!=1 or local[0].get('hidden') not in ('1','true'):
        raise Stop('УЗК: не подтверждено имя диапазона автофильтра.')
    name=local[0]
    expression=re.fullmatch(r"((?:'УЗК'|УЗК)!\$A\$1:\$C\$)([1-9][0-9]*)",name.text or '')
    if expression is None or int(expression[2])!=old_last:
        raise Stop('УЗК: имя диапазона не совпадает с границей автофильтра.')
    if any(n is not name and (n.get('localSheetId')==scope or 'УЗК!' in (n.text or '').upper()
                              or "'УЗК'!" in (n.text or '').upper()) for n in names):
        raise Stop('УЗК: проверьте дополнительные именованные диапазоны перед возвратом.')
    if new_last<=old_last:
        return False
    auto_filter.set('ref',f'A1:C{new_last}')
    name.text=expression[1]+str(new_last)
    return True


def _updated_native_filter(workbook_xml, old_sheet_xml, new_sheet_xml, removed):
    """Allow only the expected shrink of the УЗК auto-filter or its residual name."""
    import re
    from lxml import etree

    x='{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
    old_sheet=etree.fromstring(old_sheet_xml)
    new_sheet=etree.fromstring(new_sheet_xml)
    old_filter=old_sheet.find(x+'autoFilter')
    new_filter=new_sheet.find(x+'autoFilter')
    if (old_filter is None)!=(new_filter is None) or 1 in removed:
        return None

    workbook=etree.fromstring(workbook_xml)
    scope=next((str(i) for i,sheet in enumerate(workbook.find(x+'sheets'))
                if sheet.get('name')=='УЗК'),None)
    names=[node for node in workbook.findall(f'{x}definedNames/{x}definedName')
           if node.get('name')=='_xlnm._FilterDatabase' and node.get('localSheetId')==scope]
    if scope is None or len(names)!=1 or names[0].get('hidden')!='1':
        return None
    name=names[0]
    formula=re.fullmatch(r"((?:'УЗК'|УЗК)!\$A\$1:\$C\$)([1-9][0-9]*)",name.text or '')
    if formula is None:
        return None
    end=int(formula[2])
    new_end=end-bisect_left(removed,end+1)
    if end>1048576 or new_end==end or new_end<1:
        return None
    if old_filter is None:
        # Excel still adjusts a local hidden name after the filter is turned off.
        for sheet in (old_sheet,new_sheet):
            properties=sheet.find(x+'sheetPr')
            if properties is not None and properties.get('filterMode') not in (None,'0','false'):
                return None
    else:
        if old_filter.get('ref')!=f'A1:C{end}':
            return None
        old_filter.set('ref',f'A1:C{new_end}')
        if etree.tostring(old_filter,method='c14n',exclusive=True)!=etree.tostring(
                new_filter,method='c14n',exclusive=True):
            return None
    name.text=formula[1]+str(new_end)
    return etree.tostring(workbook,encoding='UTF-8',xml_declaration=True)


def _native_drawing_candidate(source,target,original,sheet_path,updates,deletions):
    """Delete rows on a private Excel copy and accept only the verified delta."""
    from io import BytesIO
    from tempfile import TemporaryDirectory
    from zipfile import ZipFile
    from openpyxl.utils.cell import coordinate_from_string
    from lxml import etree
    from olefile import OleFileIO
    from passport_app.excel import excel_session,open_excel

    def cells(path):
        workbook=load_workbook(path,read_only=True,data_only=False,keep_links=False)
        try:
            return {sheet.title:{cell.coordinate:(cell.data_type,cell.value)
                                for row in sheet.iter_rows() for cell in row
                                if cell.value is not None} for sheet in workbook}
        finally:
            workbook.close()

    def vba_streams(data):
        with OleFileIO(BytesIO(data)) as ole:
            return {tuple(name):ole.openstream(name).read() for name in ole.listdir()}

    def workbook_semantics(data):
        root=etree.fromstring(data)
        # Excel rewrites location, collaboration metadata and window position
        # on Save As a private copy. Preserve the original bytes of all three.
        names={'absPath','revisionPtr','bookViews'}
        return [etree.tostring(node,method='c14n',exclusive=True)
                for node in root if etree.QName(node).localname not in names
                and not (etree.QName(node).localname=='AlternateContent'
                         and [etree.QName(child).localname for child in node.iter()][1:]
                         ==['Choice','absPath'])]

    before=cells(source)
    removed=sorted(deletions)
    removed_set=set(removed)
    def shifted(row):
        return row-bisect_left(removed,row)

    expected={}
    for address, value in before['УЗК'].items():
        column,row=coordinate_from_string(address)
        if row not in removed_set:
            expected[column+str(shifted(row))]=value
    for row,value in updates.items():
        expected['C'+str(shifted(row))]=('n',value)
    before['УЗК']=expected

    with TemporaryDirectory(prefix='uzk-',dir=target.parent) as folder:
        native=Path(folder)/'native.xlsm'
        verified=Path(folder)/'verified.xlsm'
        from shutil import copyfile
        copyfile(source,native)
        if sha(native)!=original:
            raise Stop('УЗК изменился при копировании для Excel.')
        app=None
        try:
            with excel_session() as app:
                book=open_excel(app,native)
                sheet=None
                try:
                    sheet=book.Worksheets('УЗК')
                    for row,left in updates.items():
                        sheet.Cells(row,3).Value2=left
                    for row in sorted(deletions,reverse=True):
                        sheet.Rows(row).Delete()
                    book.Save()
                finally:
                    sheet=None
                    book.Close(SaveChanges=False)
                    book=None
        finally:
            app=None

        if cells(native)!=before:
            raise Stop('УЗК: Excel изменил неожиданные ячейки, формулы или листы.')
        allowed={sheet_path,'xl/sharedStrings.xml'}
        with ZipFile(source) as old, ZipFile(native) as new:
            if set(old.namelist())!=set(new.namelist()) or new.testzip() is not None:
                raise Stop('УЗК: Excel изменил состав частей XLSM.')
            changed={n for n in old.namelist() if old.read(n)!=new.read(n)}
            incidental={'docProps/core.xml','xl/workbook.xml','xl/vbaProject.bin'}
            if changed-allowed-incidental:
                raise Stop('УЗК: изменились графика, элементы управления или другие части XLSM: '
                           +', '.join(sorted(changed-allowed-incidental)))
            workbook='xl/workbook.xml'
            workbook_edit=None
            if workbook in changed:
                workbook_before=old.read(workbook)
                workbook_after=workbook_semantics(new.read(workbook))
                if workbook_semantics(workbook_before)!=workbook_after:
                    workbook_edit=_updated_native_filter(workbook_before,old.read(sheet_path),
                                                        new.read(sheet_path),removed)
                    if workbook_edit is None or workbook_semantics(workbook_edit)!=workbook_after:
                        raise Stop('УЗК: Excel изменил настройки книги или именованные диапазоны.')
            vba='xl/vbaProject.bin'
            if vba in changed and vba_streams(old.read(vba))!=vba_streams(new.read(vba)):
                raise Stop('УЗК: Excel изменил содержимое VBA.')
            # Keep the original metadata and VBA. Patch the one verified
            # name alongside the worksheet when Excel shortens its range.
            replacements={name:new.read(name) for name in changed & allowed}
            if workbook_edit is not None:
                replacements[workbook]=workbook_edit
            with ZipFile(verified,'w') as out:
                for item in old.infolist():
                    name=item.filename
                    out.writestr(item,replacements[name] if name in replacements else old.read(name))
        if cells(verified)!=before or sha(source)!=original:
            raise Stop('УЗК: проверка сохранённой копии после Excel не пройдена.')
        if target.exists():
            raise Stop('УЗК: кандидат XLSM уже существует.')
        verified.rename(target)
    return {'sha256':sha(target),'deleted_rows':deletions,'updated_rows':updates,
            'sheet_xml':sheet_path,'changed_parts':sorted(replacements)}
