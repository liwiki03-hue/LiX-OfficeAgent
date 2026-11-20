import json
import os
from pathlib import Path
from types import SimpleNamespace

import pytest

from tests.test_document_conversation import Answers, RetryAnswers, send
from tests.test_floating_browser import open_floating_menu
from tests.test_paged_inbox_browser import select_reference
from tests.test_model_connections import connection, payload, save, select
from tests.test_office_canvas_browser import native_browser
from tests.test_paged_browser import paged_browser
from tests.test_professional_browser import EDGE, STATIC


def test_final_answer_is_measured_when_it_replaces_an_unpainted_preview():
    if not EDGE.is_file():
        pytest.skip('Requires existing Windows Edge')
    from playwright.sync_api import sync_playwright

    with sync_playwright() as playwright:
        browser = playwright.chromium.launch(executable_path=str(EDGE), headless=True, args=['--disable-background-networking'])
        try:
            page = browser.new_page()
            page.set_content('<div id="floating-scroll"><div id="floating-timeline"></div></div>')
            page.add_script_tag(path=STATIC / 'document_conversation.js')
            page.evaluate("""() => {
                const row = {id:'rapid-answer', question:'精简', document_id:'word', status:'running',
                    intent:'edit', stage:'proposal', sequence:1, created:Date.now()/1000, text:'正在核对',
                    preview:{id:'proposal', draft_revision:0, instruction_revision:1, text:'平均延迟为120 ms。'}};
                OfficeConversation.echo({id:row.id, text:row.question, document:row.document_id});
                OfficeConversation.restore({context_id:'word', turns:[row]}, {document_id:'word',
                    draft_revision:0, instruction_revision:1, proposals:[{id:'proposal',status:'pending'}]});
                OfficeConversation.merge({...row, sequence:2, status:'complete', stage:'complete', preview:null,
                    text:'平均延迟为120 ms。', answer:'平均延迟为120 ms。'});
            }""")
            page.get_by_role('status').filter(has_text='平均延迟为120 ms。').wait_for()
            page.wait_for_function('Number.isFinite(OfficeConversation.paintTimings()[0].final_answer_ms)', timeout=1500)
            point = page.evaluate('OfficeConversation.paintTimings()[0]')
            assert point['first_effective_source'] == 'final'
            assert point['request_ms'] < point['first_effective_ms'] == point['final_answer_ms']
        finally:
            browser.close()


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
    page.wait_for_function('OfficeConversation.paintTimings().some(point=>Number.isFinite(point.question_echo_ms)&&Number.isFinite(point.final_answer_ms))')
    painted = page.evaluate('OfficeConversation.paintTimings()')
    assert len(painted) == 1 and {'id', 'request_ms', 'question_echo_ms', 'final_answer_ms'} <= painted[0].keys()
    assert painted[0]['request_ms'] <= painted[0]['question_echo_ms'] <= painted[0]['final_answer_ms']
    assert painted[0]['first_effective_source'] == 'final'
    assert painted[0]['first_effective_ms'] == painted[0]['final_answer_ms']
    page.evaluate("document.getElementById('floating-timeline').replaceChildren();OfficeConversation.reset();OfficeFloating.refresh()")
    page.locator('.conversation-turn [role="status"]').filter(has_text='第二点依据').wait_for()
    assert page.locator('.conversation-turn [role="status"]').filter(has_text='第二点依据').count() == 1
    assert len(answers.calls) == 1 and not errors
    folder = os.environ.get('LIX_CONVERSATION_EVIDENCE')
    if folder:
        directory = Path(folder);directory.mkdir(parents=True, exist_ok=True)
        page.screenshot(path=str(directory/(source.suffix[1:]+'-answer.png')))
        (directory/(source.suffix[1:]+'-timing.json')).write_text(json.dumps(page.evaluate('OfficeConversation.timings()'),ensure_ascii=False,indent=2),encoding='utf-8')
        (directory/(source.suffix[1:]+'-paint-timing.json')).write_text(json.dumps(painted,ensure_ascii=False,indent=2),encoding='utf-8')


def test_word_answer_and_same_turn_recovery(paged_browser, monkeypatch):
    page, api, source, errors = paged_browser
    answers = configure(page, api, monkeypatch)
    verify_answer(page, api, source, errors, answers, '最后一段总结了什么？')


def test_word_retry_button_preserves_original_selection(paged_browser, monkeypatch):
    page, api, source, errors = paged_browser
    configure(page, api, monkeypatch)
    model = RetryAnswers('沿用原选区。')
    monkeypatch.setattr(api._connections, 'build_model', model.build_model)
    select_reference(page)
    page.locator('#floating-input').fill('所选内容是什么？')
    page.locator('#floating-send').click()
    failed = page.locator('.conversation-turn[data-state="failed"]')
    retry = failed.get_by_role('button', name='从此处恢复', exact=True)
    retry.wait_for()
    turn_id = failed.get_attribute('data-turn-id')
    retry.click()
    reply = page.locator(f'.conversation-turn[data-turn-id="{turn_id}"][data-state="complete"]')
    reply.get_by_text('正文第2段').wait_for()
    assert page.locator('.conversation-turn').count() == 1
    assert model.builds == 2 and len(model.calls) == 1
    assert '温度120 ℃' in model.calls[0] and '非目标段落不变' not in model.calls[0]
    assert not api._floating.outputs and not api._floating.access.state()['id'] and not errors


def test_lost_receipt_replays_original_reference_instead_of_current_selection(paged_browser, monkeypatch):
    page, api, source, errors = paged_browser
    answers = configure(page, api, monkeypatch)
    select_reference(page)
    page.evaluate("""() => {
        const wrap=OfficeAuth.wrapApi;
        window.retainedAttempts=[];
        OfficeAuth.wrapApi=value=>new Proxy(wrap(value),{get(target,key){
            if(key!=='conversation_begin')return target[key];
            return async(...args)=>{
                retainedAttempts.push(args);
                if(retainedAttempts.length===1)throw Error('fixture: transport lost before accept');
                return target[key](...args);
            };
        }});
    }""")
    page.locator('#floating-input').fill('所选内容是什么？')
    page.locator('#floating-send').click()
    retry = page.locator('.conversation-turn[data-state="interrupted"]').get_by_role('button', name='从此处恢复', exact=True)
    retry.wait_for()
    select_reference(page, '非目标段落不变。')
    retry.click()
    page.locator('.conversation-turn').get_by_text('原选区已失效').wait_for()
    attempts = page.evaluate('retainedAttempts')
    assert len(attempts) == 2 and attempts[0][0] == attempts[1][0]
    assert attempts[0][5] == attempts[1][5] and attempts[0][5]
    assert not answers.calls and not api._floating.outputs and not errors


def test_natural_language_consultation_uses_no_editor_submit(paged_browser, monkeypatch):
    page, api, source, errors = paged_browser
    answers = configure(page, api, monkeypatch)
    select_reference(page)
    page.evaluate("""() => {
        window.testEditorSubmissions = [];
        OfficePaged.submit = async (...args) => {
            testEditorSubmissions.push(args);
            return {ok: true};
        };
    }""")
    for phrase, count in [('这段需要修改吗？', 1), ('这段有没有必要改', 2)]:
        page.locator('#floating-input').fill(phrase)
        page.locator('#floating-send').click()
        page.wait_for_function("count => document.querySelectorAll('.conversation-turn[data-state=complete]').length === count", arg=count)
        assert not page.evaluate('testEditorSubmissions')
        assert not api._floating.access.state()['id'] and not api._floating.paged.proposals and not api._floating.outputs
    page.locator('#floating-input').fill('能帮我把这段改成两句吗？')
    page.locator('#floating-send').click()
    page.wait_for_function('testEditorSubmissions.length === 1')
    assert page.evaluate('testEditorSubmissions[0][1]') == '能帮我把这段改成两句吗？'
    assert len(answers.calls) == 2 and not errors


def test_start_new_conversation_from_limit_and_view_old_turns(paged_browser, monkeypatch):
    page, api, source, errors = paged_browser
    answers = configure(page, api, monkeypatch)
    conversation = api._floating.conversation
    seed = send(api._floating, 'archival-seed', '进度如何？')['turn']
    with conversation.lock:
        conversation.turns.update({f'prior-{index}': dict(seed, id=f'prior-{index}', question=f'旧对话第 {index} 条')
                                   for index in range(78)})
        conversation.persist()
    await_limit = page.get_by_role('button', name='开始新对话', exact=True)
    page.locator('#floating-input').fill('进度如何？')
    page.locator('#floating-send').click()
    await_limit.wait_for()
    assert len(conversation.turns) == 80
    page.locator('#floating-input').fill('停止')
    page.locator('#floating-send').click()
    page.locator('.conversation-turn').last.get_by_text('已停止后续派发').wait_for()
    page.locator('#floating-input').fill('尚未发送的草稿要求')
    await_limit.click()
    page.wait_for_function('!document.querySelector(".conversation-turn")')
    assert page.locator('#floating-input').input_value() == '尚未发送的草稿要求'
    page.locator('#floating-input').fill('整篇文档内容是什么？')
    page.locator('#floating-send').click()
    page.locator('.conversation-turn[data-state="complete"] [role="status"]').filter(has_text='第二点依据').wait_for()
    open_floating_menu(page, '工具与记录')
    page.get_by_role('button', name='历史记录', exact=True).click()
    page.get_by_role('button', name='文档任务', exact=True).click()
    page.get_by_role('button', name='查看旧对话', exact=True).click()
    history = page.get_by_role('dialog', name='旧对话', exact=True)
    history.get_by_role('button', name='旧对话 1', exact=False).click()
    history.get_by_role('heading', name='旧对话第 77 条', exact=True).wait_for()
    assert history.locator('.floating-history-turn').count() == 80
    assert len(conversation.turns) == 1 and len(answers.calls) == 1
    assert not api._floating.outputs and not errors


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


def test_canvas_retry_button_uses_original_reference(native_browser, monkeypatch):
    page, api, source, errors = native_browser
    configure(page, api, monkeypatch)
    model = RetryAnswers('沿用原选区。')
    monkeypatch.setattr(api._connections, 'build_model', model.build_model)
    if source.suffix == '.xlsx':
        page.locator('.canvas-grid button[data-target="B2"]').click()
        selected, outside, location = '120', '合并说明', '实验 · B2'
    else:
        page.get_by_role('button', name='引用片段 run-0', exact=True).click()
        selected, outside, location = '实验汇报', '温度120', '第 1 页 · 普通文字'
    if page.locator('#canvas-toggle-ai').get_attribute('aria-expanded') != 'true':
        page.locator('#canvas-toggle-ai').click()
    page.locator('#floating-input').fill('所选内容是什么？')
    page.locator('#floating-send').click()
    failed = page.locator('.conversation-turn[data-state="failed"]')
    restore = failed.get_by_role('button', name='从此处恢复', exact=True)
    restore.click()
    reply = page.locator('.conversation-turn[data-state="complete"]')
    reply.get_by_text(location).wait_for()
    assert page.locator('.conversation-turn').count() == 1 and not errors
    assert model.builds == 2 and len(model.calls) == 1 and selected in model.calls[0] and outside not in model.calls[0]
    assert not api._floating.outputs and not api._floating.access.state()['id']


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
