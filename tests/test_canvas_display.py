from datetime import date

from openpyxl import Workbook
from openpyxl.utils.datetime import CALENDAR_MAC_1904

from office_agent import office_canvas as engine


def cells(path):
    view = engine.view(path.read_bytes(), '.xlsx', 'sheet-0', 1, 1)
    return view, {cell['id']: cell for cell in view['cells']}


def test_builtin_number_date_and_formula_display_are_read_only(tmp_path):
    source = tmp_path / '显示.xlsx'
    book = Workbook()
    sheet = book.active
    sheet.title = '显示'
    sheet['A1'] = 45658
    sheet['A1'].number_format = 'yyyy-mm-dd'
    sheet['A2'] = 1234.5
    sheet['A2'].number_format = '#,##0.00'
    sheet['A3'] = 0.125
    sheet['A3'].number_format = '0.00%'
    sheet['A4'] = 2
    sheet['A4'].number_format = '0.00'
    sheet['A5'] = 2
    sheet['A5'].number_format = 'yyyy-mm-dd hh:mm:ss.000'
    sheet['A6'] = '=A2+1'
    sheet['A6'].number_format = '#,##0.00'
    sheet['A7'] = -0.5
    sheet['A7'].number_format = '0.00'
    sheet['A8'] = 0.75
    sheet['A8'].number_format = 'h:mm'
    sheet['A9'] = 60
    sheet['A9'].number_format = 'yyyy-mm-dd'
    sheet['A10'] = -1
    sheet['A10'].number_format = 'yyyy-mm-dd'
    sheet['A11'] = 1.2345
    sheet['A11'].number_format = '0.00E+00'
    sheet['A12'] = 2.345
    sheet['A12'].number_format = '#,##0.00'
    sheet['A13'] = -2.345
    sheet['A13'].number_format = '#,##0.00'
    sheet['A14'] = 1234.5
    sheet['A14'].number_format = '#,##0'
    sheet['A15'] = 1.5
    sheet['A15'].number_format = '0,00'
    book.save(source)
    original = source.read_bytes()
    view, values = cells(source)
    assert view['date_system'] == '1900'
    assert values['A1']['display']['text'] == '2025-01-01'
    assert values['A2']['display']['text'] == '1,234.50'
    assert values['A3']['display']['text'] == '12.50%'
    assert values['A4']['display']['text'] == '2.00'
    assert values['A5']['display']['kind'] == 'unsupported'
    assert values['A5']['display']['text'] == '2'
    assert values['A6']['calculation']['status'] == 'calculated'
    assert values['A6']['display']['text'] == '1,235.50'
    assert values['A7']['display']['text'] == '-0.50'
    assert values['A8']['display']['text'] == '18:00'
    assert values['A9']['display']['kind'] == 'unsupported'
    assert values['A10']['display']['kind'] == 'unsupported'
    assert values['A11']['display']['text'] == '1.23E+00'
    assert values['A12']['display']['text'] == '2.35'
    assert values['A13']['display']['text'] == '-2.35'
    assert values['A14']['display']['text'] == '1,235'
    assert values['A15']['display']['kind'] == 'unsupported'
    assert source.read_bytes() == original


def test_1904_date_system_and_custom_iso_date(tmp_path):
    source = tmp_path / '1904.xlsx'
    book = Workbook()
    book.epoch = CALENDAR_MAC_1904
    sheet = book.active
    sheet.title = '显示'
    sheet['A1'] = date(2025, 1, 1)
    sheet['A1'].number_format = 'yyyy-mm-dd'
    sheet['A2'] = date(2025, 1, 1)
    sheet['A2'].number_format = 'yyyy/m/d'
    book.save(source)
    view, values = cells(source)
    assert view['date_system'] == '1904'
    assert values['A1']['display']['text'] == '2025-01-01'
    assert values['A1']['display']['date_system'] == '1904'
    assert values['A2']['display']['text'] == '2025/01/01'
