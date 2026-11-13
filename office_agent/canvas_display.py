from datetime import datetime, timedelta
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP

from .paged_documents import xml


S = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'

_DATE_CODES = {
    14: 'yyyy-mm-dd',
    15: 'd-mmm-yy',
    16: 'd-mmm',
    17: 'mmm-yy',
    18: 'h:mm AM/PM',
    19: 'h:mm:ss AM/PM',
    20: 'h:mm',
    21: 'h:mm:ss',
    22: 'yyyy-mm-dd hh:mm',
    45: 'mm:ss',
    46: '[h]:mm:ss',
    47: 'mmss.0',
}
_NUMBER_CODES = {
    1: '0',
    2: '0.00',
    3: '#,##0',
    4: '#,##0.00',
    9: '0%',
    10: '0.00%',
    11: '0.00E+00',
}
_CUSTOM_DATE_CODES = {
    'yyyy-mm-dd',
    'yyyy/m/d',
    'm/d/yyyy',
    'mm/dd/yyyy',
    'd-mmm-yy',
    'mmm-yy',
    'h:mm',
    'h:mm:ss',
    'yyyy-mm-dd hh:mm',
}
_CUSTOM_NUMBER_CODES = set(_NUMBER_CODES.values())


def date1904(raw):
    root = xml(raw['xl/workbook.xml']) if 'xl/workbook.xml' in raw else None
    if root is None:
        return False
    node = root.find(S+'workbookPr')
    return node is not None and node.get('date1904', '0') not in {'0', 'false'}


def _numfmt(styles, cell):
    if styles is None:
        return None
    xfs = styles.find(S+'cellXfs')
    if xfs is None:
        return None
    try:
        index = int(cell.get('s', '0'))
        xf = xfs[index]
        number_id = int(xf.get('numFmtId', '0'))
    except (IndexError, TypeError, ValueError):
        return None
    if number_id == 0:
        return None
    if number_id in _DATE_CODES:
        return _DATE_CODES[number_id]
    if number_id in _NUMBER_CODES:
        return _NUMBER_CODES[number_id]
    for node in styles.findall('./'+S+'numFmts/'+S+'numFmt'):
        if node.get('numFmtId') == str(number_id):
            return node.get('formatCode')
    return None


def _number(number, code):
    if code == '0':
        value = number.quantize(Decimal(1), rounding=ROUND_HALF_UP)
        return format(value, 'f')
    if code == '0.00':
        value = number.quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)
        return format(value, '.2f')
    if code == '#,##0':
        value = number.quantize(Decimal(1), rounding=ROUND_HALF_UP)
        return format(value, ',')
    if code == '#,##0.00':
        value = number.quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)
        return format(value, ',.2f')
    if code == '0%':
        value = (number * 100).quantize(Decimal(1), rounding=ROUND_HALF_UP)
        return format(value, 'f')+'%'
    if code == '0.00%':
        value = (number * 100).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)
        return format(value, '.2f')+'%'
    if code == '0.00E+00':
        text = format(number, '.2E')
        mantissa, exponent = text.split('E')
        power = int(exponent)
        return f'{mantissa}E{"+" if power >= 0 else "-"}{abs(power):02d}'
    raise ValueError(code)


def _date(number, code, mac):
    if number < 0:
        raise ValueError('negative date serial')
    day = int(number)
    if day == 60 and not mac:
        raise ValueError('1900 leap-day serial')
    if mac:
        base = datetime(1904, 1, 1)
        moment = base + timedelta(days=day)
    else:
        base = datetime(1899, 12, 31) if day < 60 else datetime(1899, 12, 30)
        moment = base + timedelta(days=day)
    fraction = number - day
    moment += timedelta(seconds=float(fraction * Decimal(86400)))
    if code == 'yyyy-mm-dd':
        return moment.strftime('%Y-%m-%d')
    if code == 'yyyy/m/d':
        return moment.strftime('%Y/%m/%d')
    if code == 'm/d/yyyy':
        return f'{moment.month}/{moment.day}/{moment.year}'
    if code == 'mm/dd/yyyy':
        return moment.strftime('%m/%d/%Y')
    if code == 'd-mmm-yy':
        return moment.strftime('%d-%b-%y')
    if code == 'd-mmm':
        return moment.strftime('%d-%b')
    if code == 'mmm-yy':
        return moment.strftime('%b-%y')
    if code == 'h:mm AM/PM':
        return moment.strftime('%I:%M %p')
    if code == 'h:mm:ss AM/PM':
        return moment.strftime('%I:%M:%S %p')
    if code == 'h:mm':
        return moment.strftime('%H:%M')
    if code == 'h:mm:ss':
        return moment.strftime('%H:%M:%S')
    if code == 'yyyy-mm-dd hh:mm':
        return moment.strftime('%Y-%m-%d %H:%M')
    if code == 'mm:ss':
        return moment.strftime('%M:%S')
    if code == '[h]:mm:ss':
        hours = int(number * 24)
        remainder = moment.replace(second=0, microsecond=0)
        return f'{hours:02d}:{remainder.strftime("%M:%S")}'
    if code == 'mmss.0':
        return moment.strftime('%M%S')+'.'+str(moment.microsecond // 100000)
    raise ValueError(code)


def describe(styles, cell, value, kind, mac):
    if kind != 'n' or value in {None, ''}:
        return None
    code = _numfmt(styles, cell)
    if not code:
        return None
    try:
        number = Decimal(str(value))
        if not number.is_finite():
            raise InvalidOperation
        normalized = code.strip()
        if normalized in _DATE_CODES.values() or normalized in _CUSTOM_DATE_CODES:
            return {'kind': 'date', 'text': _date(number, normalized, mac), 'format': normalized,
                    'date_system': '1904' if mac else '1900', 'approximate': True}
        if normalized in _NUMBER_CODES.values() or normalized in _CUSTOM_NUMBER_CODES:
            return {'kind': 'number', 'text': _number(number, normalized), 'format': normalized}
        return {'kind': 'unsupported', 'text': str(value), 'format': normalized,
                'reason': '当前仅显示内置数值/日期格式和明确白名单格式'}
    except (InvalidOperation, OverflowError, ValueError):
        return {'kind': 'unsupported', 'text': str(value), 'format': code,
                'reason': '数值或日期超出当前只读显示边界'}
