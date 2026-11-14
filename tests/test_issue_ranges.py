import base64
import io
import json
from pathlib import Path
from types import SimpleNamespace
import threading

from docx import Document
from docx.text.paragraph import Paragraph
from openpyxl import Workbook, load_workbook
import pytest

from office_agent.canvas_ranges import read_range, used_selection
from office_agent.paged_documents import W, GuardError
from office_agent.webui.bridge import WebviewApp
from office_agent.webui.floating_session import FloatingSession
from tests.test_canvas_ai import canvas, finish
from tests.test_office_canvas import wait_finished


def test_frozen_cross_window_range_reads_actual_values_and_header_budget():
    book = Workbook(); sheet = book.active; sheet.freeze_panes = 'D3'
    for row in range(1, 66):
        for col in range(1, 22):
            sheet.cell(row, col, f'{row}:{col}')
    stream = io.BytesIO(); book.save(stream); data = stream.getvalue()
    region = read_range(data, 'sheet-0', 'O28', 'R35')
    assert region['rows'] == 8 and region['cols'] == 4
    assert all(cell['value'] == f'{row}:{col}' for row, values in enumerate(region['cells'], 28) for col, cell in enumerate(values, 15))
    assert used_selection(data, 'sheet-0', 'row', 35)['end'] == 'U35'
    assert used_selection(data, 'sheet-0', 'column', 18)['end'] == 'R65'
    sheet['A101'] = '预算外'; stream = io.BytesIO(); book.save(stream)
    with pytest.raises(GuardError, match='100'):
        used_selection(stream.getvalue(), 'sheet-0', 'column', 1)


@pytest.mark.parametrize('canvas', ['.xlsx'], indirect=True)
def test_multi_cell_ai_atomic_copy_and_formula_protection(canvas, monkeypatch):
    service, floating, source, _, _ = canvas
    original = source.read_bytes()
    reference = service.ai.select_range(service.document_id, 0, 'sheet-0', 'A1', 'C2', ['B2'])['reference']
    cells = {cell['target']: cell for cell in reference['cells']}
    assert not cells['C2']['editable'] and cells['C2']['formula'] == 'B2*2'
    assert cells['B2']['editable']
    from tests.test_canvas_ai import Connections
    connections = Connections(service, '')
    def model(route):
        class Model:
            def bind_tools(self, tools):
                assert tools[0]['function']['name'] == 'propose_canvas_range'
                return self
            def invoke(self, messages):
                if any(isinstance(row, dict) and row.get('role') == 'tool' for row in messages):
                    return {'content': '两个普通目标已形成建议。'}
                return {'tool_calls': [{'name': 'propose_canvas_range', 'id': 'range-call', 'args': {'reference_id': reference['id'],
                    'edits': [{'target': 'A1', 'old': '温度', 'new': '温度说明'}, {'target': 'B2', 'old': '120', 'new': '130'}], 'reason': '仅修改明确目标'}}]}
        return Model()
    connections.build_model = model; floating.api._connections = connections
    service.ai.submit(service.document_id, 0, reference['id'], '', 'range-first', '修改选定文字和已允许的B2数值，公式保持。')
    finish(service.ai)
    proposal = next(iter(service.ai.proposals.values()))
    assert len(proposal['operations']) == 2
    service.ai.adopt(service.document_id, 0, proposal['id'])
    assert service.revision == 1
    result = load_workbook(io.BytesIO(service.data))['实验']
    assert result['A1'].value == '温度说明' and result['B2'].value == 130 and result['C2'].value == '=B2*2'
    monkeypatch.setattr('office_agent.webui.office_canvas.compare_documents', lambda *a, **k: {'inputs_unchanged': True, 'render': {'status': 'different'}})
    plan = service.plan(service.document_id, 1)['plan']; service.publish(service.document_id, 1, plan['id'], True); wait_finished(service)
    assert Path(plan['output']).is_file() and source.read_bytes() == original
    assert load_workbook(plan['output'])['实验']['B2'].value == 130
    report = Path(__file__).parents[1] / 'deliverables/issue-repair-20261114/Q4_RANGE_COPY.json'
    report.parent.mkdir(parents=True, exist_ok=True)
    report.write_text(json.dumps({
        'edits': 2, 'reopened': True, 'source_unchanged': True, 'formula_unchanged': True, 'render': 'controlled_test_double', 'provider': 'simulated'}, indent=2), 'utf8')


@pytest.mark.parametrize('canvas', ['.xlsx'], indirect=True)
def test_range_rejects_ungranted_numeric_formula_and_partial_batch(canvas):
    from office_agent.webui.canvas_ai import CanvasCellChange
    service, floating, source, _, _ = canvas
    reference = service.ai.select_range(service.document_id, 0, 'sheet-0', 'A1', 'C2')['reference']
    before = service.data
    for target, old in [('B2', '120'), ('C2', '')]:
        with pytest.raises(GuardError):
            service.ai._range_operations(reference, [CanvasCellChange(target='A1', old='温度', new='说明'), CanvasCellChange(target=target, old=old, new='2')])
        assert service.data == before and service.revision == 0
    merged = service.ai.select_range(service.document_id, 0, 'sheet-0', 'A4', 'C4')['reference']
    assert all(not row['editable'] and row['merged'] for row in merged['cells'])


@pytest.mark.parametrize('canvas', ['.xlsx'], indirect=True)
def test_new_range_and_inflight_requirement_preserve_session_and_block_old_action(canvas):
    from tests.test_canvas_ai import Connections
    service, floating, source, _, _ = canvas
    reference = service.ai.select_range(service.document_id, 0, 'sheet-0', 'A1', 'C2')['reference']
    entered, release = threading.Event(), threading.Event()
    connections = Connections(service, '')
    count = []
    def model(route):
        selected = service.ai.reference.copy()
        class Model:
            def bind_tools(self, tools): return self
            def invoke(self, messages):
                count.append(1)
                if len(count) == 1:
                    entered.set(); assert release.wait(10)
                if any(isinstance(row, dict) and row.get('role') == 'tool' for row in messages): return {'content': '建议已形成'}
                return {'tool_calls': [{'name': 'propose_canvas_range', 'id': 'bounded-range', 'args': {'reference_id': selected['id'],
                    'edits': [{'target': 'A1', 'old': '温度', 'new': '温度说明'}], 'reason': '仅所选普通文字'}}]}
        return Model()
    connections.build_model = model; floating.api._connections = connections
    service.ai.submit(service.document_id, 0, reference['id'], '', 'one', '补充温度说明')
    assert entered.wait(5)
    runtime = service.ai.runtime
    try:
        latest = service.ai.select_range(service.document_id, 0, 'sheet-0', 'A1', 'B1')['reference']
        assert service.ai.runtime is runtime and not runtime.closed
        service.ai.submit(service.document_id, 0, latest['id'], runtime.session_id, 'two', '保留其余目标与数值')
    finally:
        release.set(); finish(service.ai)
    assert len(service.ai.proposals) == 1 and len(count) == 3
    proposal = next(iter(service.ai.proposals.values()))
    assert proposal['reference_id'] == latest['id'] and proposal['instruction_revision'] == 2
    assert service.ai.runtime is runtime and len(service.ai.state()['requirements']) == 2 and service.revision == 0


def test_word_table_reference_uses_current_nodes_utf16_and_complex_readonly(tmp_path):
    source = tmp_path / '表格.docx'; book = Document(); table = book.add_table(rows=2, cols=2)
    table.cell(0, 0).text = '说明😀'; table.cell(0, 1).text = '目标文字'; table.cell(1, 0).text = '120 ℃'; table.cell(1, 1).text = '保持'
    book.save(source); original = source.read_bytes()
    app = WebviewApp(); api = SimpleNamespace(_app=app, _auth=SimpleNamespace(status=lambda: {'authenticated': True}), _memory_key=lambda: 'isolated', _connections=SimpleNamespace(state=SimpleNamespace(mode='agent')))
    floating = FloatingSession(api); assert floating.admit(str(source))['ok']; assert floating.read(floating.read_token, True)['ok']
    service = floating.paged; service.open(floating.document['id'], floating.document['version'])
    def sync(book, revision):
        stream = io.BytesIO(); book.save(stream)
        manifest = [{'id': f'p-{index}', 'text': Paragraph(node, book).text} for index, node in enumerate(book.element.body.iter(W + 'p'))]
        if revision: service.changed(service.document_id, revision)
        service.sync(service.document_id, revision, base64.b64encode(stream.getvalue()).decode(), manifest, revision == 0)
        return manifest
    manifest = sync(book, 0)
    nodes = [{'node_id': row['id'], 'start': 0, 'end': len(row['text'].encode('utf-16-le')) // 2} for row in manifest]
    reference = service.reference(service.document_id, 0, {'nodes': nodes, 'whole_table': True})['reference']
    assert len(reference['cells']) == 4 and all(row['editable'] for row in reference['cells'])
    assert reference['cells'][0]['end'] == 4 and reference['cells'][0]['old'] == '说明😀'
    assert len({row['cell_id'] for row in reference['cells']}) == 4
    with pytest.raises(GuardError, match='整表'):
        service.reference(service.document_id, 0, {'nodes': nodes[:1], 'whole_table': True})
    book.tables[0].cell(0, 0).merge(book.tables[0].cell(0, 1)); manifest = sync(book, 1)
    first = manifest[0]
    reference = service.reference(service.document_id, 1, {'nodes': [{'node_id': first['id'], 'start': 0, 'end': len(first['text'].encode('utf-16-le')) // 2}], 'whole_table': False})['reference']
    assert not reference['cells'][0]['editable'] and source.read_bytes() == original
