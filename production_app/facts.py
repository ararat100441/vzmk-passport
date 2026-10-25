"""Human confirmed source data, never simulated by production algorithms."""
from datetime import date,datetime
from decimal import Decimal,InvalidOperation
from pathlib import Path
import re,zipfile
from .storage import Stop,safe,sha

def numeric(value,label,minimum=None,maximum=None):
    try:v=Decimal(str(value).replace(',','.'))
    except InvalidOperation:raise Stop('Требуется число: '+label)
    if not v.is_finite() or minimum is not None and v<minimum or maximum is not None and v>maximum:raise Stop('Недопустимое значение: '+label)
    return v

def validate_layer_ranges(layers):
    """Exactly three positive per-layer ranges, in micrometres.

    Ranges are reusable specifications/preferences, not evidence of measured
    dry/wet coating thickness on a newly manufactured structure.
    """
    if not isinstance(layers, list) or len(layers) != 3:
        raise Stop('Для трёхслойного АКЗ задайте диапазоны толщины всех трёх слоёв.')
    result = {}
    for entry in layers:
        if not isinstance(entry, dict) or type(entry.get('layer')) is not int:
            raise Stop('Неверный номер слоя АКЗ.')
        n = entry['layer']
        if n not in (1, 2, 3) or n in result:
            raise Stop('Слои АКЗ должны иметь уникальные номера 1, 2, 3.')
        low = numeric(entry.get('min'), 'минимальная толщина слоя '+str(n), Decimal('0.001'))
        high = numeric(entry.get('max'), 'максимальная толщина слоя '+str(n), Decimal('0.001'))
        if low > high:
            raise Stop('Минимальная толщина превышает максимальную: слой '+str(n))
        result[n] = {'layer': n, 'min': format(low, 'f'), 'max': format(high, 'f')}
    return [result[n] for n in (1, 2, 3)]



LAYER_DETAILS_REQUIRED=('material','paint_batch','hardener_batch','color',
                        'application_date','time_start','time_end','wet_average')


def validate_layer_details(entries, facts):
    """Confirm three distinct per-coat real facts for the historical АКЗ act.

    Daily climate is shared only when all coats are actually applied on the
    same confirmed day. Application time, batch, paint and wet-film mean are
    independent facts for each coat; never reuse old passport measurements.
    """
    if not isinstance(entries,list) or len(entries)!=3:
        raise Stop('Для трёхслойного акта АКЗ требуются сведения каждого из трёх слоёв.')
    indexed={}
    start=date.fromisoformat(facts['start_date'])
    end=date.fromisoformat(facts['date'])
    akz_day=date.fromisoformat(facts['akz_date'])
    for item in entries:
        if not isinstance(item,dict) or type(item.get('layer')) is not int:
            raise Stop('Некорректный номер слоя АКЗ.')
        number=item['layer']
        if number not in (1,2,3) or number in indexed:
            raise Stop('Слои АКЗ должны быть 1, 2, 3 без повторов.')
        payload={'layer':number}
        for key in LAYER_DETAILS_REQUIRED:
            value=item.get(key)
            if not isinstance(value,str) or not value.strip() or '{{' in value or '}}' in value:
                raise Stop('Не подтверждено '+key+' слоя '+str(number))
            payload[key]=value.strip()
        try:actual_day=date.fromisoformat(payload['application_date'])
        except ValueError as exc:raise Stop('Дата нанесения слоя АКЗ должна быть YYYY-MM-DD.') from exc
        if not start<=actual_day<=end or actual_day!=akz_day:
            raise Stop('Один климат на все слои допустим только при совпадении даты нанесения с подтверждённой датой АКЗ.')
        import re
        def minutes(text):
            if not re.fullmatch(r'\d{2}:\d{2}',text):raise Stop('Время нанесения должно иметь формат HH:MM.')
            h,m=map(int,text.split(':'))
            if h>23 or m>59:raise Stop('Некорректное время нанесения АКЗ.')
            return 60*h+m
        if minutes(payload['time_end'])<=minutes(payload['time_start']):
            raise Stop('Время окончания слоя АКЗ должно быть позднее начала в тот же день.')
        numeric(payload['wet_average'],'Средняя мокрая толщина '+str(number),Decimal('0.001'))
        indexed[number]=payload
    return [indexed[n] for n in (1,2,3)]


def validate_dry_averages(layers, averages):
    """Compare measured mean dry coating after each layer with cumulative limits.

    The three measurements are cumulative as in the supplied issued Excel act:
    average after coat 1, average after coats 1+2, and average after 1+2+3.
    Ranges are allowed specifications, not measured min/max observations.
    """
    limits = validate_layer_ranges(layers)
    if not isinstance(averages, list) or len(averages) != 3:
        raise Stop('Укажите среднюю толщину сухого покрытия после каждого из трёх слоёв.')
    raw = {}
    for item in averages:
        if not isinstance(item, dict) or type(item.get('layer')) is not int:
            raise Stop('Неверный номер этапа измерения средней толщины АКЗ.')
        n = item['layer']
        if n not in (1, 2, 3) or n in raw:
            raise Stop('Средние толщины АКЗ должны соответствовать слоям 1, 2 и 3 без повторов.')
        raw[n] = numeric(item.get('average'), 'средняя сухая толщина после слоя '+str(n), Decimal('0.001'))
    low = high = previous = Decimal(0)
    results = []
    for layer in limits:
        n = layer['layer']
        low += Decimal(layer['min'])
        high += Decimal(layer['max'])
        measured = raw[n]
        if measured < previous:
            raise Stop('Средняя сухая толщина после слоя '+str(n)+' меньше предыдущего этапа.')
        if not low <= measured <= high:
            raise Stop('Средняя сухая толщина после слоя '+str(n)+
                       ' выходит за суммарный диапазон '+format(low,'f')+'–'+format(high,'f')+' мкм.')
        results.append({'layer':n, 'average':format(measured,'f'),
                        'required_total_min':format(low,'f'),
                        'required_total_max':format(high,'f')})
        previous = measured
    return results


def validate(facts,vom,project,number,training=False,require_review=False,*,defer_registry_dates=False,defer_approvals=False):
    if facts.get('project')!=project or facts.get('number')!=number:raise Stop('Результаты контроля относятся к другому объекту/паспорту.')
    if not defer_approvals and (facts.get('confirmed') is not True or facts.get('authorized') is not True):
        raise Stop('Нет подтверждения фактических результатов и права на выпуск.')
    # For the agreed simple interface, no personal name is requested.
    # Old templates still contain RESPONSiBLE field: it is left blank,
    # never populated with a person from the historical first passport.
    if not facts.get('no_responsible_input'):
        if not isinstance(facts.get('responsible'),str) or not facts['responsible'].strip():
            raise Stop('Не указано ответственное лицо.')
    if facts.get('kind')!='ACTUAL' and not training:raise Stop('TEST/SAMPLE_ONLY/учебные данные запрещены для рабочего выпуска.')
    if facts.get('kind') not in ('ACTUAL','TRAINING'):raise Stop('Не указан тип результатов контроля.')
    if not defer_approvals and facts.get('forms_approved') is not True:raise Stop('Ответственное лицо не утвердило формы документов.')
    if not defer_registry_dates:
        for field in ('start_date','date','vik_date','uzk_date','akz_date'):
            try:date.fromisoformat(facts[field])
            except (KeyError,ValueError,TypeError):
                raise Stop('Не подтверждена дата в реестре: '+field)
        start,end=date.fromisoformat(facts['start_date']),date.fromisoformat(facts['date'])
        if start>end or any(not start<=date.fromisoformat(facts[k])<=end
                            for k in ('vik_date','uzk_date','akz_date')):
            raise Stop('Даты контроля вне подтверждённого периода изготовления.')
    if require_review and not facts.get('documents_reviewed'):raise Stop('Документы ещё не просмотрены ответственным лицом.')
    expected={}
    for row in vom.rows:expected[str(row[0])]=expected.get(str(row[0]),Decimal(0))+row[2]
    rows=facts.get('marks',[])
    if len({str(r.get('mark')) for r in rows})!=len(rows) or {str(r.get('mark')) for r in rows}!=set(expected):raise Stop('Марки фактического контроля не совпадают с видимой ВОМ.')
    variants=set()
    for r in rows:
        mark=str(r['mark']);q=numeric(r.get('uzk_quantity'),mark,0,expected[mark])
        if not training:
            if r.get('vik_result')!='ГОДЕН' or not str(r.get('vik_basis','')).strip():
                raise Stop('Нет подтверждённого результата ВИК и первичного протокола: '+mark)
        if q!=q.to_integral_value():raise Stop('Количество УЗК должно быть целым: '+mark)
        if r.get('uzk_result') not in ('ГОДЕН','НЕ ТРЕБУЕТСЯ'):raise Stop('Нет подтверждённого результата УЗК: '+mark)
        if r['uzk_result']=='НЕ ТРЕБУЕТСЯ' and (q!=0 or not str(r.get('basis','')).strip()):raise Stop('Укажите основание отсутствия УЗК: '+mark)
        if r['uzk_result']=='ГОДЕН' and (q==0 or not str(r.get('profile','')).strip()):raise Stop('Нет профиля/количества УЗК: '+mark)
        variant=r.get('akz')
        if variant in (None,''):
            if not str(r.get('basis','')).strip():raise Stop('Не подтверждено отсутствие АКЗ: '+mark)
        elif type(variant)!=int or variant not in range(4):raise Stop('Некорректный № АКЗ: '+mark)
        else:
            variants.add(variant)
            if not str(r.get('paint','')).strip():raise Stop('Не подтверждена краска: '+mark)
        if not training and not str(r.get('basis','')).strip():
            raise Stop('Нет первичного протокола/основания УЗК: '+mark)
    from .real_forms import normalize_doc_fields,validate_uzk_details
    normalize_doc_fields(facts.get('doc_fields'))
    if not defer_registry_dates:
        validate_uzk_details(facts,vom)
    climates=facts.get('climates',[])
    if {r.get('variant') for r in climates}!=variants or len(climates)!=len(variants):raise Stop('Нет отдельного подтверждённого набора измерений для каждого выбранного АКЗ.')
    for c in climates:
        air=numeric(c.get('air'),'Температура воздуха',-60,80)
        surface=numeric(c.get('surface'),'Температура поверхности',-60,100)
        dew=numeric(c.get('dew'),'Точка росы',-80,80)
        numeric(c.get('humidity'),'Влажность',0,100)
        if c.get('layers'):
            # For three coats, the official act records MEASURED mean dry
            # thickness after each successive coat, not measured min/max.
            validate_dry_averages(c['layers'],c.get('dry_averages'))
        else:
            numeric(c.get('thickness'),'Средняя толщина',Decimal('0.001'))
        from .measurement_ranges import (
            verify_midrange_provenance,validate_temperature_order)
        verify_midrange_provenance(c)
        validate_temperature_order(c)
        if c.get('layer_details') and not defer_registry_dates:
            validate_layer_details(c['layer_details'],facts)
        if not str(c.get('material','')).strip() or not str(c.get('basis','')).strip():raise Stop('Нет материала/протокола послойных измерений АКЗ.')
    proofs=facts.get('evidence',[])
    if not proofs:raise Stop('Приложите первичные протоколы контроля с подтверждёнными результатами.')
    for proof in proofs:
        path=safe(proof['path'])
        if not path.is_file() or sha(path)!=proof['sha256']:raise Stop('Изменился/отсутствует первичный протокол: '+str(path))
        if not training:
            suffix=path.suffix.lower()
            if suffix in ('.docx','.xlsx'):
                with zipfile.ZipFile(path) as z:text=' '.join(z.read(n).decode('utf8',errors='replace') for n in z.namelist() if n.endswith('.xml'))
            elif suffix in ('.txt','.csv'):text=path.read_text(encoding='utf-8-sig')
            elif suffix=='.pdf':
                from pypdf import PdfReader
                text=' '.join(p.extract_text() or '' for p in PdfReader(path).pages)
            else:raise Stop('Протокол нельзя проверить автоматически. Нужен читаемый PDF/DOCX/XLSX/TXT: '+str(path))
            if not text.strip():raise Stop('Протокол не содержит читаемого текста: '+str(path))
            if re.search(r'\b(TEST|TRAINING|SAMPLE_ONLY|ТЕСТ|ЧЕРНОВИК)\b|УЧЕБН|RANDBETWEEN|СЛУЧМЕЖДУ',text,re.I):raise Stop('В первичном протоколе обнаружены TEST/учебные/модельные данные: '+str(path))
    return sorted(variants)
