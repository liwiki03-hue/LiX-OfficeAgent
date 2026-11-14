import io
import json
from pathlib import Path
import time

from docx import Document
import pytest

from tests.test_office_canvas_browser import native_browser
from tests.test_paged_browser import paged_browser, engine_frame


def test_requirements_queue_without_valid_selection_or_model(paged_browser, monkeypatch):
    page, api, source, errors = paged_browser
    original = source.read_bytes()
    calls = []
    monkeypatch.setattr(api._connections, 'build_model', lambda route: calls.append(route))
    page.locator('#floating-input').fill('先保存这条要求，引用后再精简文字，保留数字与单位。')
    page.locator('#floating-send').click()
    page.wait_for_function("document.getElementById('floating-input').value === ''")
    runtime = api._floating.runtime
    assert runtime and runtime.revision == 1 and len(runtime.messages) == 1
    assert api._floating.paged.state()['pending_reason'] == 'reference'
    assert not calls and not api._floating.paged.references and not api._floating.outputs
    assert source.read_bytes() == original and not errors


@pytest.mark.parametrize('native_browser', ['.xlsx'], indirect=True)
def test_pointer_header_keyboard_and_context_reference(native_browser):
    page, api, source, errors = native_browser
    start = page.locator('.canvas-grid button[data-target="A1"]').bounding_box()
    end = page.locator('.canvas-grid button[data-target="C2"]').bounding_box()
    page.mouse.move(start['x'] + 15, start['y'] + 12); page.mouse.down()
    page.mouse.move(end['x'] + 15, end['y'] + 12, steps=8); page.mouse.up()
    assert page.locator('#canvas-position').text_content() == 'A1:C2'
    assert page.locator('.canvas-grid [aria-selected="true"]').count() >= 6
    page.locator('.canvas-grid button[data-target="C2"]').press('Shift+ArrowDown')
    page.wait_for_function("document.querySelector('#canvas-position').textContent==='A1:C3'")
    page.get_by_role('button', name='选择已用范围内整行 1', exact=True).click()
    page.wait_for_function("document.querySelector('#canvas-position').textContent==='A1:C1'")
    page.locator('.canvas-grid button[data-target="A1"]').click(button='right')
    menu = page.get_by_role('dialog', name='选区 A1:C1', exact=True)
    menu.get_by_role('button', name='引用给 AI', exact=True).click()
    page.wait_for_function("document.querySelector('#canvas-ai-panel').textContent.includes('A1:C1')")
    reference = api._floating.canvas.ai.reference
    assert reference['kind'] == 'range' and len(reference['cells']) == 3
    assert reference['cells'][0]['cell_id'] == 'sheet-0!A1'
    page.locator('.canvas-grid button[data-target="A1"]').dblclick()
    assert page.locator('#canvas-value').evaluate('node=>node===document.activeElement')
    assert not errors
    out = Path(__file__).parents[1] / 'deliverables/issue-repair-completion-20261114'
    out.mkdir(parents=True, exist_ok=True)
    page.screenshot(path=str(out / 'Q4-grid-selection.png'))


def test_word_entire_table_ai_batch_actual_editor_copy_and_reopen(paged_browser, monkeypatch):
    page, api, source, errors = paged_browser
    floating = api._floating; service = floating.paged; original = source.read_bytes()
    floating.access.set(floating.document['id'], floating.document['version'], floating.epoch, 'edit', True)
    page.locator('#simple-model-menu > summary').click()
    page.locator('#model-selector').select_option(json.dumps(['official', 'same-model', 'agent'], separators=(',', ':'))); page.keyboard.press('Escape')
    frame = engine_frame(page)
    frame.locator('.superdoc-page').get_by_text('条件', exact=True).first.click()
    page.get_by_label('引用整表，最多8个段落', exact=True).click()
    assert page.locator('#paged-table-reference .agent-icon').evaluate('node=>getComputedStyle(node).maskImage') != 'none'
    page.wait_for_function("document.querySelector('#paged-reference-state').textContent.startsWith('已引用')")
    reference = next(iter(service.references.values()))
    assert reference['kind'] == 'table' and len(reference['cells']) == 2
    class Model:
        def bind_tools(self, tools):
            assert tools[0]['function']['name'] == 'propose_draft_table_edit'
            return self
        def invoke(self, messages):
            if any(isinstance(row, dict) and row.get('role') == 'tool' for row in messages):
                return {'content': '普通表格文字建议已形成。'}
            return {'tool_calls': [{'name': 'propose_draft_table_edit', 'id': 'table-change', 'args': {'reference_id': reference['id'], 'reason': '补充说明，保留数字单位和表格结构',
                'edits': [{'node_id': row['node_id'], 'old': row['old'], 'new': '实验条件' if row['old'] == '条件' else '120 ms（复核）'} for row in reference['cells']]}}]}
    monkeypatch.setattr(api._connections, 'build_model', lambda route: Model())
    page.locator('#floating-input').fill('补充表格文字说明，保留120 ms与表格结构。'); page.locator('#floating-send').click()
    deadline = time.monotonic() + 25
    while not service.proposals or not all(row['status'] == 'adopted' for row in service.proposals.values()):
        if time.monotonic() > deadline: raise AssertionError({'feedback': page.locator('#paged-feedback').text_content(), 'state': service.state()})
        page.wait_for_timeout(50)
    page.wait_for_timeout(300)
    book = Document(io.BytesIO(service.data))
    assert book.tables[0].cell(0, 0).text == '实验条件' and book.tables[0].cell(1, 1).text == '120 ms（复核）'
    assert len(book.tables[0].rows) == 2 and len(book.tables[0].columns) == 2
    page.get_by_role('button', name='保存副本', exact=True).click()
    page.get_by_role('button', name='确认生成副本', exact=True).click()
    deadline = time.monotonic() + 150
    while not floating.outputs:
        if time.monotonic() > deadline: raise AssertionError(page.locator('#paged-feedback').text_content())
        page.wait_for_timeout(100)
    receipt = next(iter(floating.outputs.values())); output = Path(receipt['output'])
    reopened = Document(output)
    assert reopened.tables[0].cell(0, 0).text == '实验条件' and reopened.tables[0].cell(1, 1).text == '120 ms（复核）'
    assert source.read_bytes() == original and not errors
    out = Path(__file__).parents[1] / 'deliverables/issue-repair-completion-20261114'
    out.mkdir(parents=True, exist_ok=True)
    page.screenshot(path=str(out / 'Q4-word-table-copy.png'))
    (out / 'Q4_WORD_BROWSER.json').write_text(json.dumps({'real_editor': 'SuperDoc1.47.2/headless Edge', 'edits': 2, 'copy_reopened': True,
        'source_unchanged': True, 'structure_preserved': True, 'provider': 'simulated', 'render': receipt.get('render_diff', {}).get('render', {}).get('status')}, ensure_ascii=False, indent=2), 'utf8')
