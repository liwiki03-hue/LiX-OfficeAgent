import io

from openpyxl import load_workbook
import pytest

from office_agent.canvas_formulas import Calculator
from office_agent.office_canvas import S, xml
from office_agent.writes import GuardError
from tests.test_office_canvas import canvas, operation


@pytest.mark.parametrize('canvas', ['.xlsx'], indirect=True)
def test_sheet_cache_isolation_and_unseen_sheet_edit(canvas):
    service, floating, source, _, _ = canvas
    doc = floating.document['id']
    service.apply(doc, 0, operation('B2', '7', '', 'sheet-1', 'number'))
    service.apply(doc, 1, operation('C2', '=B2*3', '', 'sheet-1', 'formula'))
    first = service.view(doc, 2, 'sheet-0')['view']
    second = service.view(doc, 2, 'sheet-1')['view']
    result = lambda snapshot: next(cell for cell in snapshot['cells'] if cell['id'] == 'C2')['calculation']['value']
    assert result(first) == '240'
    assert result(second) == '21'
    assert result(service.view(doc, 2, 'sheet-0')['view']) == '240'
    service.apply(doc, 2, operation('B2', '8', '7', 'sheet-1', 'number'))
    assert result(service.view(doc, 3, 'sheet-0')['view']) == '240'
    assert result(service.view(doc, 3, 'sheet-1')['view']) == '24'
    service.undo(doc, 3)
    assert result(service.view(doc, 4, 'sheet-1')['view']) == '21'


@pytest.mark.parametrize('formula, expected', [('=MOD(-7,3)', '2'), ('=MOD(7,-3)', '-2'), ('=MOD(-7,-3)', '-1')])
def test_mod_divisor_sign(formula, expected):
    sheet = xml(('<worksheet xmlns="'+S[1:-1]+'"><sheetData/></worksheet>').encode())
    assert Calculator(sheet).calculate('A1', formula)['value'] == expected


def test_sumproduct_rejects_equal_count_different_shapes():
    sheet = xml(('<worksheet xmlns="'+S[1:-1]+'"><sheetData><row r="1"><c r="A1"><v>1</v></c><c r="B1"><v>2</v></c></row><row r="2"><c r="A2"><v>3</v></c></row></sheetData></worksheet>').encode())
    with pytest.raises(GuardError, match='形状|行列'):
        Calculator(sheet).calculate('C1', '=SUMPRODUCT(A1:B1,A1:A2)')


@pytest.mark.parametrize('formula, expected, kind', [
    ('=LEN("中文 Ab")', '5', 'number'),
    ('=TRIM("  保留  条件  ")', '保留 条件 ', 'text'),
    ('=CONCATENATE("aBc","中文","""引号""")', 'aBc中文"引号"', 'text'),
    ('=IF(0,"先前","继续")', '继续', 'text'),
    ('=NUMBERVALUE("1.234,56",",",".")', '1234.56', 'number'),
    ('=NUMBERVALUE("-1,234.56",".",",")', '-1234.56', 'number'),
    ('=NUMBERVALUE("12 345,67",","," ")', '12345.67', 'number'),
    ('=NUMBERVALUE("2,5%",",",".")', '0.025', 'number'),
    ('=ROUND(NUMBERVALUE("-2,345",",","."),2)', '-2.35', 'number'),
])
def test_typed_text_and_explicit_decimal_conversion(formula, expected, kind):
    sheet = xml(('<worksheet xmlns="'+S[1:-1]+'"><sheetData/></worksheet>').encode())
    result = Calculator(sheet).calculate('A1', formula)
    assert result['value'] == expected and result['value_type'] == kind


@pytest.mark.parametrize('formula', [
    '=NUMBERVALUE("1,25")', '=NUMBERVALUE("",".",",")', '=NUMBERVALUE("1.2,3",".",",")',
    '=NUMBERVALUE("1,2",",",",")', '=NUMBERVALUE("１２,５",",",".")', '=NUMBERVALUE("NaN",".",",")',
    '=NUMBERVALUE("1234567890123456",".",",")', '=SUM("1",2)', '=TRIM(1)', '=LEN("🙂")',
    '=CONCATENATE(A1:B1)', '=NUMBERVALUE("1e5",".",",")', '=1,25+1',
])
def test_ambiguous_or_unsupported_text_numeric_inputs_fail_closed(formula):
    sheet = xml(('<worksheet xmlns="'+S[1:-1]+'"><sheetData/></worksheet>').encode())
    result = Calculator(sheet).describe('A1', formula)
    assert result['status'] == 'unavailable' and result['reason']


@pytest.mark.parametrize('canvas', ['.xlsx'], indirect=True)
def test_text_formula_plan_replay_and_cache_dependencies(canvas):
    service, floating, source, _, _ = canvas
    doc = floating.document['id']; original = source.read_bytes()
    service.apply(doc, 0, operation('D2', '=CONCATENATE(A1," · 记录")', '', action='formula'))
    for _ in range(2):
        result = service.view(doc, 1, 'sheet-0')['view']
        cell = next(cell for cell in result['cells'] if cell['id'] == 'D2')
        assert cell['calculation']['value'] == '温度 · 记录'
        assert cell['display']['text'] == '温度 · 记录'
        assert cell['calculation']['depends_on'] == ['A1']
        assert cell['formula_editable'] and not cell['editable']
    service.apply(doc, 1, operation('A1', '测试', '温度'))
    final = service.plan(doc, 2)['plan']
    assert final['changes'][0]['calculation']['value'] == '测试 · 记录'
    assert load_workbook(io.BytesIO(service.data))['实验']['D2'].value == '=CONCATENATE(A1," · 记录")'
    assert source.read_bytes() == original
    assert service.publish(doc, 2, final['id'], False)['cancelled']


def test_shared_text_reference_and_missing_value_are_distinct():
    sheet = xml(('<worksheet xmlns="'+S[1:-1]+'"><sheetData><row r="1"><c r="A1" t="s"><v>0</v></c></row></sheetData></worksheet>').encode())
    text = xml(('<si xmlns="'+S[1:-1]+'"><t>  Ab中文 </t></si>').encode())
    calculator = Calculator(sheet, shared=[text])
    assert calculator.calculate('B1', '=TRIM(A1)')['value'] == 'Ab中文'
    assert calculator.describe('B1', '=TRIM(A2)')['status'] == 'unavailable'
    assert calculator.describe('B1', '=SUM(A1)')['status'] == 'unavailable'
    assert Calculator(sheet).describe('B1', '=TRIM(A1)')['status'] == 'unavailable'


@pytest.mark.parametrize('canvas', ['.xlsx'], indirect=True)
def test_prefetch_is_single_bounded_entry_rechecked_before_consumption(canvas, monkeypatch):
    from office_agent import office_canvas as engine
    service, floating, source, _, _ = canvas; doc = floating.document['id']
    calls = []
    original_view = engine.view
    def counted(*args, **kwargs):
        calls.append(args[2:5]); return original_view(*args, **kwargs)
    monkeypatch.setattr(engine, 'view', counted)
    service.prefetch(doc, 0, 'sheet-0', 11, 1)
    service.prefetch(doc, 0, 'sheet-0', 11, 1)
    snapshot = service.view(doc, 0, 'sheet-0', 11, 1)['view']
    assert len(calls) == 1 and snapshot['row'] == 11
    assert len(snapshot['row_indices']) <= 30 and len(snapshot['col_indices']) <= 16
    assert service.window_cache is None
    service.prefetch(doc, 0, 'sheet-0', 21, 1)
    service.prefetch(doc, 0, 'sheet-1', 31, 1)
    assert service.window_cache[0] == (doc, 0, 'sheet-1', 31, 1)
    service.apply(doc, 0, operation('B2', '125', '120', action='number'))
    assert service.window_cache is None
    with pytest.raises(GuardError): service.view(doc, 0, 'sheet-1', 31, 1)
    service.prefetch(doc, 1, 'sheet-0', 11, 1)
    with pytest.raises(GuardError): service.view(doc, 1, 'sheet-0', 11, True)
    source.write_bytes(source.read_bytes()+b'changed')
    with pytest.raises((ValueError, GuardError)): service.view(doc, 1, 'sheet-0', 11, 1)


@pytest.mark.parametrize('canvas', ['.xlsx'], indirect=True)
def test_cache_invalidation_tracks_dependencies_for_each_sheet(canvas):
    service, floating, _, _, _ = canvas; doc = floating.document['id']
    service.apply(doc, 0, operation('D2', '=C2+1', '', action='formula'))
    service.view(doc, 1, 'sheet-0')
    service.view(doc, 1, 'sheet-1')
    service.apply(doc, 1, operation('B2', '125', '120', action='number'))
    snapshot = service.view(doc, 2, 'sheet-0')['view']
    assert next(cell for cell in snapshot['cells'] if cell['id'] == 'D2')['calculation']['value'] == '251'
