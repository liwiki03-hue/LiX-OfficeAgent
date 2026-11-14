from copy import deepcopy
import json
import os
from pathlib import Path
import shutil

from docx import Document
import pytest

from tests.test_paged_inbox_browser import paged_browser, select_model, select_reference, wait


@pytest.mark.parametrize('question', [
    '把年份改为2026，保持温度和数字不变。',
    '只将选中句中的2025年改为2026年，保留温度120 ℃、标点及其他文字，不修改标题和表格。',
])
def test_explicit_year_proposal_cancel_publish_reopen_and_rollback(paged_browser, monkeypatch, question):
    page, api, source, errors = paged_browser
    original = source.read_bytes()
    requests = []
    class Model:
        def __init__(self):
            self.reference = deepcopy(api._floating.paged.references[api._floating.paged.pending_reference_id])
            self.calls = 0
        def bind_tools(self, tools):
            return self
        def invoke(self, messages):
            self.calls += 1
            requests.append(self.calls)
            if self.calls == 1:
                return {'tool_calls': [{'name':'propose_draft_edit','id':'year-case','args': {
                    'reference_id':self.reference['id'], 'old':self.reference['old'],
                    'new':self.reference['old'].replace('2025','2026'), 'reason':'将明确指定的年份改为2026，保留温度和单位。'}}]}
            return {'content':'年份修改建议已生成，尚未写入文件。'}
    monkeypatch.setattr(api._connections, 'build_model', lambda route: Model())
    select_model(page)
    select_reference(page)
    page.locator('#floating-input').fill(question)
    page.locator('#floating-send').click()
    wait(page,lambda:any(row['status']=='adopted' for row in api._floating.paged.proposals.values()))
    assert '2025 → 2026' in page.locator('#paged-proposals').inner_text()
    assert 'same-model' in page.locator('#paged-proposals').inner_text()
    wait(page, lambda: not api._app.busy)
    wait(page,lambda:any(row['status']=='adopted' for row in api._floating.paged.proposals.values()))
    wait(page, lambda: any(row['status']=='adopted' for row in api._floating.paged.proposals.values()))
    page.get_by_role('button',name='保存副本',exact=True).click()
    page.get_by_role('dialog').get_by_role('button',name='取消',exact=True).click()
    assert source.read_bytes() == original and not list(source.parent.glob('*.agent-*.docx'))
    page.get_by_role('button',name='保存副本',exact=True).click()
    page.get_by_role('button',name='确认生成副本',exact=True).click()
    wait(page, lambda: bool(api._floating.outputs), 180)
    receipt = next(iter(api._floating.outputs.values()))
    output = Path(receipt['output'])
    book = Document(output)
    assert book.paragraphs[1].text == '实验于2026年完成，温度120 ℃。'
    assert book.paragraphs[2].text == '非目标段落不变。'
    assert book.tables[0].cell(1,1).text == '120 ms'
    assert book.sections[0].header.paragraphs[0].text == '自建验收页眉'
    assert book.sections[0].footer.paragraphs[0].text == '自建验收页脚'
    assert receipt['status']=='verified' and receipt['render_diff']['inputs_unchanged'] and requests==[1,2]
    page.wait_for_function("document.getElementById('paged-save-state').textContent.includes('副本已验证')")
    folder = os.getenv('OFFICE_AGENT_INBOX_EVIDENCE')
    evidence = Path(folder)/'explicit-year' if folder else None
    if evidence:
        evidence.mkdir(parents=True, exist_ok=True)
        shutil.copy2(source, evidence/'source.docx')
        shutil.copy2(output, evidence/'verified-copy.docx')
        page.screenshot(path=str(evidence/'verified-copy.png'))
    api._app.editor.rollback(receipt['id'], confirmed=True)
    assert not output.exists() and source.read_bytes()==original and errors==[]
    if evidence:
        (evidence/'RESULT.json').write_text(json.dumps({'scope':'headless Edge and real SuperDoc; controlled model; real copy/render/reopen/rollback; no native Office or service',
            'model_calls':len(requests), 'source_unchanged':True, 'rollback':True, 'render':receipt['render_diff']},ensure_ascii=False,indent=2),encoding='utf8')


def test_document_history_reentry_refreshes_even_when_navigation_remained_history(paged_browser, monkeypatch):
    page, api, source, errors = paged_browser
    page.wait_for_function('!historyLoading')
    listings = []
    def latest():
        listings.append(True)
        return {'ok': True, 'sessions': []}
    monkeypatch.setattr(api, 'list_sessions', latest)
    page.locator('#floating-input').fill('文档草稿不能丢失')
    page.evaluate('window.OfficeFloating.history()')
    page.wait_for_function('!historyLoading')
    assert len(listings) == 1
    page.locator('#floating-return').click()
    assert page.locator('#floating-input').input_value() == '文档草稿不能丢失'
    page.evaluate('window.OfficeFloating.history()')
    page.wait_for_function('!historyLoading')
    assert len(listings) == 2
    assert errors == []
