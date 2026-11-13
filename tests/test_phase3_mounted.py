import json
from pathlib import Path

from docx import Document
from openpyxl import Workbook, load_workbook
import pytest

from office_agent.auth import LocalAuth
from office_agent.webui import bridge
from office_agent.writes import Editor, digest


@pytest.fixture
def mounted(tmp_path):
    docs = tmp_path/'documents'; docs.mkdir()
    word = docs/'one.docx'; document = Document(); document.add_paragraph('Original'); document.save(word)
    excel = docs/'two.xlsx'; book = Workbook(); book.active['A1'] = 1; book.save(excel)
    auth = LocalAuth(tmp_path/'application')
    assert auth.register('testuser','Testing123')['ok']
    assert auth.login('testuser','Testing123')['ok']
    app = bridge.WebviewApp(); app.path = word; app.editor = Editor(docs,require_confirmation=True)
    api = bridge.Api(app,auth=auth)
    return api,app,word,excel


def test_general_memory_api_is_separate_and_injected_only_after_confirmation(mounted):
    api,app,word,_ = mounted
    assert api.memory_snapshot()['result']['enabled'] is False
    content = '长期任务是整理季度汇总'
    plan = api.plan_general_memory('remember',{'content':content,'category':'task'})['result']
    assert api.general_memory_snapshot()['result']['items'] == []
    assert api.confirm_general_memory(plan['plan_id'],True)['result']['status'] == 'saved'
    questions = []
    app.ask = lambda question:questions.append(question) or {'ok':True}
    assert api.ask('整理季度汇总')['ok']
    assert app._memory_context[0]['text'] == content
    assert questions == ['整理季度汇总']
    assert api.memory_snapshot()['result']['enabled'] is False
    assert api.logout()['ok']
    assert api.general_memory_snapshot()['reason'] == 'not_authenticated'
    assert app._memory_context is None


def test_multi_scope_api_has_no_group_execute(mounted):
    api,app,word,excel = mounted
    plan = api.plan_task_documents([word.name,excel.name])['result']
    assert app._multi_scope is None
    assert not api.confirm_task_documents(plan['plan_id'],False)['ok']
    plan = api.plan_task_documents([word.name,excel.name])['result']
    assert api.confirm_task_documents(plan['plan_id'],True)['ok']
    scope = app._multi_scope
    result = scope.dispatcher()('plan_word',{'path':str(word),'paragraph_index':0,'old':'Original','new':'Updated'})
    listing = api.task_documents_snapshot()['result']
    assert listing['files'] == [word.name,excel.name] and len(listing['plans']) == 1
    assert api.choose_document_plan(result['plan_id'])['ok']
    assert app.pending_id == result['plan_id']
    assert not list(word.parent.glob('*.agent-*'))
    assert api.clear_task_documents()['ok']
    assert app._multi_scope is None and app.pending_id is None and app.editor.plans == {}


def test_multi_execute_refuses_change_in_other_selected_file(mounted):
    api,app,word,excel = mounted
    plan = api.plan_task_documents([word.name,excel.name])['result']
    assert api.confirm_task_documents(plan['plan_id'],True)['ok']
    result = app._multi_scope.dispatcher()('plan_word',{'path':str(word),'paragraph_index':0,'old':'Original','new':'Updated'})
    assert api.choose_document_plan(result['plan_id'])['ok']
    book = load_workbook(excel); book.active['A1'] = 99; book.save(excel)
    assert app.execute()['reason'] == 'multi_scope_changed'
    assert not list(word.parent.glob('*.agent-*'))


def test_stream_is_incremental_not_persisted_and_final_answer_once(mounted,monkeypatch):
    api,app,word,_ = mounted
    class Model:
        def bind_tools(self,schemas): return self
        def stream(self,messages):
            yield {'content':'第一部分'}
            assert any(event.get('type') == 'answer_draft_delta' for event in app._buffer)
            yield {'content':'第二部分','response_metadata':{'finish_reason':'stop'}}
    monkeypatch.setattr(app,'_build_model',lambda:Model())
    app._run_task(word,app.editor,'读取并概述文档')
    events = app.poll()
    assert [event['text'] for event in events if event['type'] == 'answer_draft_delta'] == ['第一部分','第二部分']
    assert [event['text'] for event in events if event['type'] == 'answer'] == ['第一部分第二部分']
    histories = list(app.editor.audit.glob('sessions/*/events.json'))
    assert len(histories) == 1
    history = json.loads(histories[0].read_text(encoding='utf-8'))
    assert not any(event['type'].startswith('answer_draft') for event in history['events'])
    assert len([event for event in history['events'] if event['type'] == 'answer']) == 1


def test_logout_during_stream_drops_late_text(mounted,monkeypatch):
    api,app,word,_ = mounted
    editor = app.editor
    class Model:
        def bind_tools(self,schemas): return self
        def stream(self,messages):
            yield {'content':'draft-before-logout'}
            api.logout()
            yield {'content':'late-secret-must-not-display'}
    monkeypatch.setattr(app,'_build_model',lambda:Model())
    app._run_task(word,editor,'读取并概述文档')
    assert app.path is None
    assert not any('late-secret' in str(event) for event in app._buffer)
    assert not any(event.get('type') == 'answer' for event in app._buffer)


def test_multi_task_actual_agent_uses_all_exact_paths(mounted,monkeypatch):
    api,app,word,excel = mounted
    plan = api.plan_task_documents([word.name,excel.name])['result']
    assert api.confirm_task_documents(plan['plan_id'],True)['ok']
    replies = iter([
        {'tool_calls':[{'name':'request_write_tools','args':{}}]},
        {'tool_calls':[{'name':'plan_word','args':{'path':str(word),'paragraph_index':0,'old':'Original','new':'Updated'}}]},
        {'tool_calls':[{'name':'request_read_tools','args':{'pack':'spreadsheet'}}]},
        {'tool_calls':[{'name':'request_write_tools','args':{}}]},
        {'tool_calls':[{'name':'plan_excel','args':{'path':str(excel),'sheet':'Sheet','cell':'A1','value':5}}]},
        {'content':'两个计划等待各自确认'}])
    visible = []
    class Model:
        def bind_tools(self,schemas):
            visible.append({schema['function']['name'] for schema in schemas})
            return self
        def invoke(self,messages): return next(replies)
    monkeypatch.setattr(app,'_build_model',lambda:Model())
    app._run_task(word,app.editor,'为两个已选文档分别生成修改计划',clarified=True)
    assert len(app.editor.plans) == 2
    assert {plan.request.path for plan in app.editor.plans.values()} == {str(word),str(excel)}
    assert not list(word.parent.glob('*.agent-*'))
    assert 'plan_word' not in visible[0] and 'plan_excel' not in visible[0]
    assert all(tool.allowed_paths == (word,excel) for tool in app.model_agent.capability_tools.values())


def test_new_apis_refuse_busy_and_unauthenticated(mounted):
    api,app,word,excel = mounted
    app.busy = True
    assert api.general_memory_snapshot()['reason'] == 'busy'
    assert api.plan_task_documents([word.name,excel.name])['reason'] == 'busy'
    app.busy = False
    api.logout()
    assert api.plan_general_memory('remember',{})['reason'] == 'not_authenticated'
    assert api.task_documents_snapshot()['reason'] == 'not_authenticated'
