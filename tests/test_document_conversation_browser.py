import json
import os
from pathlib import Path
from types import SimpleNamespace

from tests.test_document_conversation import Answers
from tests.test_model_connections import connection, payload, save, select
from tests.test_office_canvas_browser import native_browser
from tests.test_paged_browser import paged_browser


def configure(page, api, monkeypatch):
    configured = connection('read-test')
    configured['models'] = [{'id':'simulated', 'capabilities':['agent']}]
    save(api._connections, payload([configured]))
    select(api._connections, 'read-test', 'agent', 'simulated')
    answers = Answers()
    monkeypatch.setattr(api._connections, 'build_model', answers.build_model)
    page.evaluate('window.OfficeModelSelector.refresh()')
    return answers


def verify_answer(page, api, source, errors, answers, question):
    original = source.read_bytes()
    revision = api._floating.paged.revision if source.suffix == '.docx' else api._floating.canvas.revision
    page.locator('#floating-input').fill(question)
    page.locator('#floating-send').click()
    page.locator('.conversation-turn[data-state="complete"] [role="status"]').filter(has_text='第二点依据').wait_for(timeout=15000)
    assert page.locator('.floating-message.user p').filter(has_text=question).count() == 1
    assert page.locator('.conversation-turn [role="status"]').filter(has_text='第二点依据').count() == 1
    assert not page.locator('dialog[open]').count()
    assert not api._floating.outputs and not api._floating.access.state()['id']
    assert source.read_bytes() == original and len(answers.calls) == 1
    assert (api._floating.paged.revision if source.suffix == '.docx' else api._floating.canvas.revision) == revision
    page.evaluate("document.getElementById('floating-timeline').replaceChildren();OfficeConversation.reset();OfficeFloating.refresh()")
    page.locator('.conversation-turn [role="status"]').filter(has_text='第二点依据').wait_for()
    assert page.locator('.conversation-turn [role="status"]').filter(has_text='第二点依据').count() == 1
    assert len(answers.calls) == 1 and not errors
    folder = os.environ.get('LIX_CONVERSATION_EVIDENCE')
    if folder:
        directory = Path(folder);directory.mkdir(parents=True, exist_ok=True)
        page.screenshot(path=str(directory/(source.suffix[1:]+'-answer.png')))
        (directory/(source.suffix[1:]+'-timing.json')).write_text(json.dumps(page.evaluate('OfficeConversation.timings()'),ensure_ascii=False,indent=2),encoding='utf-8')


def test_word_answer_and_same_turn_recovery(paged_browser, monkeypatch):
    page, api, source, errors = paged_browser
    answers = configure(page, api, monkeypatch)
    verify_answer(page, api, source, errors, answers, '最后一段总结了什么？')


def test_canvas_answer_numeric_read_without_write_permission(native_browser, monkeypatch):
    page, api, source, errors = native_browser
    answers = configure(page, api, monkeypatch)
    if source.suffix == '.xlsx':
        page.locator('.canvas-grid button[data-target="B2"]').click()
    else:
        page.get_by_role('button', name='引用片段 run-0', exact=True).click()
    if page.locator('#canvas-toggle-ai').get_attribute('aria-expanded') != 'true':
        page.locator('#canvas-toggle-ai').click()
    verify_answer(page, api, source, errors, answers, '这段内容表示什么？')


def test_retry_local_message_keeps_new_unsent_draft_and_original_id(paged_browser, monkeypatch):
    page, api, source, errors = paged_browser
    answers = configure(page, api, monkeypatch)
    page.evaluate("""() => {
      const wrap=OfficeAuth.wrapApi;
      window.__lostCalls=[];
      OfficeAuth.wrapApi=value=>new Proxy(wrap(value),{get(target,key){
        if(key==='conversation_begin')return async(...args)=>{
          window.__lostCalls.push(args);
          if(window.__lostCalls.length===1)throw Error('simulated lost receipt');
          return target[key](...args);
        };
        return target[key];
      }});
    }""")
    question = '最后一段总结了什么？'
    draft = '另一条尚未发送的问题'
    page.locator('#floating-input').fill(question)
    page.locator('#floating-send').click()
    retry = page.locator('.conversation-turn[data-state="interrupted"] [data-retry-turn]')
    retry.wait_for()
    page.locator('#floating-input').fill(draft)
    retry.click()
    page.locator('.conversation-turn[data-state="complete"] [role="status"]').filter(has_text='第二点依据').wait_for(timeout=15000)
    calls = page.evaluate('window.__lostCalls')
    assert len(calls) == 2 and calls[0][0] == calls[1][0] and calls[1][1] == question
    assert page.locator('#floating-input').input_value() == draft
    assert page.locator('.floating-message.user p').filter(has_text=question).count() == 1
    assert len(answers.calls) == 1 and not api._floating.outputs and not errors


def test_late_completed_chat_status_cannot_replace_new_warning(tmp_path, monkeypatch):
    from tests.test_daily_browser import setup_daily
    from tests.test_professional_browser import _browser
    from tests.test_ui_simplify_browser import simplify, pick_model
    fixture = _browser(tmp_path, monkeypatch)
    try:
        page, api, original, calls, receipts = next(fixture)
        save(api._connections, payload([connection()]))
        setup_daily(page, api, original, calls, receipts, monkeypatch)
        simplify(page)
        pick_model(page, 'chat', 'same-model')
        monkeypatch.setattr(api._connections, 'build_model', lambda route: SimpleNamespace(invoke=lambda messages: SimpleNamespace(content='保留研究条件。', response_metadata={'finish_reason':'length'})))
        page.locator('#prompt').fill('说明摘要要求')
        page.locator('#btn-ask').click()
        page.wait_for_function('document.getElementById("chat-notice").textContent.includes("尚未完整")')
        page.wait_for_function('document.getElementById("status-card").hidden')
        book = api.conversation_snapshot()['result']
        session = next(row for row in book['sessions'] if row['id'] == book['current'])
        page.evaluate("""event=>{
          setStatus('warn','新的操作结果不确定，请先核对，不要重复提交。');
          handleEvent(event);
        }""", {'type':'status', 'channel':'chat', 'session_id':session['id'], 'turn_id':session['turns'][-1]['id'],
               'level':'warn', 'text':'回复达到输出预算，可能不完整；可核对后主动续写。'})
        page.wait_for_function('!document.getElementById("status-card").hidden')
        assert '结果不确定' in page.locator('#error-line').inner_text()
        assert calls.count('ask') == 1
    finally:
        fixture.close()
