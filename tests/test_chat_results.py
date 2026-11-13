import json
import io
from pathlib import Path
from types import SimpleNamespace

import httpx2
import pytest
from PIL import Image

from tests.test_daily_flows import setup, finish
from tests.test_model_connections import connection, payload, save, select


def setting(api, values):
    plan = api.plan_settings_update(values)
    assert plan['ok'], plan
    result = api.confirm_settings_update(plan['plan']['plan_id'], True)
    assert result['ok'], result


@pytest.mark.parametrize('field', ['max_tokens', 'max_completion_tokens'])
def test_output_budget_wire_matches_preview(tmp_path, monkeypatch, field):
    app, api, _ = setup(tmp_path, True)
    row = connection()
    row['models'][0]['output_parameter'] = field
    save(api._connections, payload([row])); select(api._connections, 'official')
    setting(api, {'chat_reply_tokens': 2048})
    calls = []
    def handle(request):
        calls.append(json.loads(request.content))
        return httpx2.Response(200, json={'id':'fixture','object':'chat.completion','created':1,'model':'same-model', 'choices':[{'index':0,'message':{'role':'assistant','content':'long'},'finish_reason':'length'}], 'usage':{'prompt_tokens':10,'completion_tokens':20,'total_tokens':30}})
    monkeypatch.setattr('office_agent.webui.model_connections._http_client', lambda: httpx2.Client(transport=httpx2.MockTransport(handle)))
    preview = api.preview_chat_context('question')['result']
    assert preview['reply_token_reserve'] == 2048 and preview['output_parameter'] == field
    assert api.ask('question', {'stream':False, 'context_revision':preview['context_revision']})['ok']; finish(app)
    assert len(calls) == 1 and calls[0][field] == 2048
    assert ('max_tokens' if field == 'max_completion_tokens' else 'max_completion_tokens') not in calls[0]
    book = api.conversation_snapshot()['result']
    turn = book['sessions'][0]['turns'][0]
    assert turn['budget']['reply_token_reserve'] == 2048 and turn['usage']['output_tokens'] == 20
    assert json.loads(Path(book['location']).read_text(encoding='utf8'))['version'] == 2
    assert api.reload_conversations()['ok']


def test_continuation_binding_and_budget_are_not_optional(tmp_path, monkeypatch):
    app, api, _ = setup(tmp_path, True)
    calls = []
    def invoke(messages):
        calls.append(messages)
        return SimpleNamespace(content='answer part',response_metadata={'finish_reason':'length'})
    monkeypatch.setattr(api._connections, 'build_model', lambda route:SimpleNamespace(invoke=invoke))
    assert api.ask('original question', {'stream':False})['ok']; finish(app)
    state = api.conversation_snapshot()['result']
    binding = {'session_id':state['current'], **state['continuations'][0]}
    question = 'continue'
    preview = api.preview_chat_context(question, {'continuation':binding})['result']
    assert preview['continuation_required'] and preview['context_turns'] == 1
    assert api.ask(question, {'stream':False,'continuation':binding,'context_revision':preview['context_revision']})['ok']; finish(app)
    assert [item['content'] for item in calls[1]] == ['original question','answer part','continue']
    latest = api.conversation_snapshot()['result']['sessions'][0]['turns']
    assert len(latest) == 2 and latest[1]['continuation_of'] == binding
    assert api.reload_conversations()['ok']
    changed = {**binding,'source_digest':'0'*64}
    assert api.ask(question, {'continuation':changed})['reason'] == 'chat_continuation_changed'
    setting(api, {'chat_reply_tokens':1024})
    assert api.ask(question, {'continuation':binding,'context_revision':preview['context_revision']})['reason'] == 'chat_context_changed'
    assert len(calls) == 2


def test_partial_stream_error_preserves_received_content(tmp_path, monkeypatch):
    app, api, _ = setup(tmp_path, True)
    def stream(messages):
        yield SimpleNamespace(content='received', response_metadata={})
        raise TimeoutError()
    monkeypatch.setattr(api._connections,'build_model', lambda route:SimpleNamespace(stream=stream))
    assert api.ask('question')['ok']; finish(app)
    turn = api.conversation_snapshot()['result']['sessions'][0]['turns'][0]
    assert turn['status'] == 'failed' and turn['answer'] == 'received'


def test_local_limit_preserves_content_and_filtered_has_no_continuation(tmp_path, monkeypatch):
    app, api, _ = setup(tmp_path, True)
    monkeypatch.setattr(api._connections,'build_model', lambda route:SimpleNamespace(invoke=lambda messages:SimpleNamespace(content='字'*32001,response_metadata={'finish_reason':'stop'})))
    assert api.ask('question',{'stream':False})['ok']; finish(app)
    state = api.conversation_snapshot()['result']
    turn = state['sessions'][0]['turns'][0]
    assert turn['finish_reason'] == 'local_limit' and len(turn['answer']) == 32000 and state['continuations']
    assert api.reload_conversations()['ok']


@pytest.mark.parametrize('kind', ['md','txt'])
def test_export_host_snapshot_selection_cancel_confirm_and_undo(tmp_path, monkeypatch, kind):
    app, api, _ = setup(tmp_path, True)
    monkeypatch.setattr(api._connections,'build_model',lambda route:SimpleNamespace(invoke=lambda messages:SimpleNamespace(content='# 原始回答\n\n**保持语法**',response_metadata={'finish_reason':'length'})))
    for text in ['first question','not selected']:
        assert api.ask(text,{'stream':False})['ok']; finish(app)
    row = api.conversation_snapshot()['result']['sessions'][0]
    app.window = SimpleNamespace(create_file_dialog=lambda *args:[str(tmp_path)])
    payload = {'session_id':row['id'],'turn_ids':[row['turns'][0]['id']],'format':kind,'metadata':False,'content':'answer'}
    plan = api.plan_chat_export(payload); assert plan['ok'],plan
    target = Path(plan['plan']['path']); assert not target.exists()
    assert api.confirm_chat_export(plan['plan']['plan_id'],False)['result']['wrote_nothing']
    assert not target.exists()
    plan = api.plan_chat_export(payload); assert plan['ok']
    saved = api.confirm_chat_export(plan['plan']['plan_id'],True)['result']
    assert saved['status'] == 'saved' and saved['verified']
    raw = Path(saved['path']).read_text(encoding='utf8')
    assert '# 原始回答\n\n**保持语法**' in raw and 'first question' not in raw and 'not selected' not in raw
    assert '达到输出预算' in raw and 'fixture-private-' not in raw and '接口：' not in raw
    assert api.confirm_chat_export(plan['plan']['plan_id'],True)['result']['status'] == 'rejected'
    assert api.plan_chat_export_rollback(saved['save_id'])['ok']
    assert api.confirm_chat_export_rollback(saved['save_id'],True)['result']['status'] == 'rolled_back'
    assert not Path(saved['path']).exists()


def test_export_conflict_logout_and_unknown_result_never_retry(tmp_path,monkeypatch):
    app,api,active = setup(tmp_path)
    monkeypatch.setattr(api._connections,'build_model',lambda route:SimpleNamespace(invoke=lambda messages:SimpleNamespace(content='ok')))
    assert api.ask('q',{'stream':False})['ok']; finish(app)
    row = api.conversation_snapshot()['result']['sessions'][0]
    app.window = SimpleNamespace(create_file_dialog=lambda *args:[str(tmp_path)])
    data = {'session_id':row['id'],'turn_ids':[row['turns'][0]['id']],'format':'txt','metadata':False}
    plan = api.plan_chat_export(data)['plan']
    monkeypatch.setattr(api._chat_exports.entries[plan['plan_id']]['archive'],'_exclusive',lambda *args:(_ for _ in ()).throw(PermissionError()))
    result = api.confirm_chat_export(plan['plan_id'],True)['result']
    assert result['status'] == 'uncertain' and result['applied'] is None
    assert api.confirm_chat_export(plan['plan_id'],True)['result']['status'] == 'rejected'
    active[0] = False
    assert api.plan_chat_export(data)['reason'] == 'not_authenticated'


def saved_image(tmp_path,monkeypatch,reuse=True):
    app,api,active = setup(tmp_path)
    if reuse:
        setting(api,{'image_reuse_metadata':True})
    select(api._connections,'official','image','image-model')
    assert api.ask('画一张蓝色科研插图',{'size':'1024x1536','style':'vivid'})['ok']
    _,service = app._image_session
    identifier = next(iter(service._plans))
    raw = io.BytesIO(); Image.new('RGB',(32,32),'blue').save(raw,format='PNG')
    calls=[]
    service._provider_factory=lambda:SimpleNamespace(generate=lambda *args:calls.append(args) or raw.getvalue())
    assert api.confirm_image_generation(identifier,True)['ok']
    app.window=SimpleNamespace(create_file_dialog=lambda *args,**kwargs:[str(tmp_path)])
    plan=api.plan_image_save(identifier)['result']
    result=api.confirm_image_save(plan['save_id'],True)
    assert result['ok'],result
    assert api.list_independent_image_history()['ok']
    return app,api,active,identifier,Path(result['result']['path']),calls


def test_historical_image_preview_reuse_and_open_are_read_only(tmp_path,monkeypatch):
    app,api,active,identifier,path,calls=saved_image(tmp_path,monkeypatch)
    before=path.read_bytes()
    assert len(calls)==1
    preview=api.preview_historical_image(identifier)
    assert preview['ok'] and preview['image']['data_url'].startswith('data:image/png;base64,'), preview
    draft=api.reuse_historical_image(identifier)
    assert draft['ok'] and draft['draft']=={'prompt':'画一张蓝色科研插图','size':'1024x1536','style':'vivid','n':1}
    opened=[]; monkeypatch.setattr('os.startfile',lambda value:opened.append(value))
    assert api.open_historical_image_folder(identifier)['ok'] and opened==[str(tmp_path)]
    assert path.read_bytes()==before and len(calls)==1
    active[0]=False
    assert api.preview_historical_image(identifier)['reason']=='not_authenticated'
    assert len(calls)==1


@pytest.mark.parametrize('fault',['missing','replaced','large','corrupt','audit','hardlink'])
def test_historical_image_rejects_changed_or_unsafe_files_without_generation(tmp_path,monkeypatch,fault):
    app,api,_,identifier,path,calls=saved_image(tmp_path,monkeypatch,False)
    if fault=='missing': path.unlink()
    elif fault=='replaced':
        data=path.read_bytes(); path.unlink(); path.write_bytes(data)
    elif fault=='large':
        with path.open('r+b') as stream: stream.truncate(10000001)
    elif fault=='corrupt': path.write_bytes(b'invalid')
    elif fault=='audit':
        _,audit,_=api._connections.image_tickets[identifier]
        body=json.loads(audit.read_bytes()); body['saved_output']['path']=str(tmp_path/'other.png'); audit.write_text(json.dumps(body),encoding='utf8')
    else:
        import os
        os.link(path,tmp_path/'linked.png')
    result=api.preview_historical_image(identifier)
    assert not result['ok'],result
    assert len(calls)==1
    assert not api.open_historical_image_folder(identifier)['ok']


def test_legacy_image_prompt_unavailable_and_restarted_preview_no_old_authority(tmp_path,monkeypatch):
    from office_agent.webui.bridge import Api,WebviewApp
    app,api,_,identifier,path,calls=saved_image(tmp_path,monkeypatch,False)
    assert api.reuse_historical_image(identifier)['reason']=='image_history_prompt_unavailable'
    restored=Api(WebviewApp(),auth=api._auth,settings_root=tmp_path)
    assert restored.preview_historical_image(identifier)['reason']=='image_history_expired'
    assert restored.list_independent_image_history()['ok']
    assert restored.preview_historical_image(identifier)['ok']
    assert restored.confirm_image_generation(identifier,True)['reason']=='no_session'
    assert len(calls)==1


def test_continuation_overflow_does_not_drop_parent_and_cross_connection_requires_confirmation(tmp_path,monkeypatch):
    app,api,_=setup(tmp_path)
    monkeypatch.setattr(api._connections,'build_model',lambda route:SimpleNamespace(invoke=lambda messages:SimpleNamespace(content='长回答 '*1800,response_metadata={'finish_reason':'length'})))
    assert api.ask('最初问题',{'stream':False})['ok']; finish(app)
    state=api.conversation_snapshot()['result']; binding={'session_id':state['current'],**state['continuations'][0]}
    row=connection('custom'); row['context_window']=1024
    save(api._connections,payload([connection(),row])); select(api._connections,'custom')
    blocked=api.preview_chat_context('继续',{'continuation':binding})
    assert blocked['reason']=='model_context_budget' and blocked['budget']['continuation_required']
    roomy=connection('custom'); roomy['context_window']=32768
    save(api._connections,payload([connection(),roomy])); select(api._connections,'custom')
    result=api.ask('继续',{'stream':False,'continuation':binding})
    assert result['reason']=='chat_transfer_required' and result['route']['connection_id']=='custom'
    assert not app.busy


def test_filtered_response_and_export_path_conflict(tmp_path,monkeypatch):
    app,api,_=setup(tmp_path)
    monkeypatch.setattr(api._connections,'build_model',lambda route:SimpleNamespace(invoke=lambda messages:SimpleNamespace(content='过滤说明',response_metadata={'finish_reason':'content_filter'})))
    assert api.ask('q',{'stream':False})['ok']; finish(app)
    state=api.conversation_snapshot()['result']; row=state['sessions'][0]
    assert state['continuations']==[]
    app.window=SimpleNamespace(create_file_dialog=lambda *args:[str(tmp_path)])
    plan=api.plan_chat_export({'session_id':row['id'],'turn_ids':[row['turns'][0]['id']],'format':'txt','metadata':True})['plan']
    Path(plan['path']).parent.mkdir()
    result=api.confirm_chat_export(plan['plan_id'],True)['result']
    assert result['wrote_nothing'] and result['code']=='chat_export_path_conflict'


def test_explicit_partial_export_is_fixed_while_stream_continues(tmp_path,monkeypatch):
    import threading
    app,api,_=setup(tmp_path)
    ready,release=threading.Event(),threading.Event()
    def stream(messages):
        yield SimpleNamespace(content='第一部分',response_metadata={})
        ready.set()
        assert release.wait(5)
        yield SimpleNamespace(content='后续部分',response_metadata={'finish_reason':'stop'})
    monkeypatch.setattr(api._connections,'build_model',lambda route:SimpleNamespace(stream=stream))
    assert api.ask('question')['ok'] and ready.wait(5)
    try:
        state=api.conversation_snapshot()['result']; row=state['sessions'][0]
        app.window=SimpleNamespace(create_file_dialog=lambda *args:[str(tmp_path)])
        data={'session_id':row['id'],'turn_ids':[row['turns'][0]['id']],'format':'md','metadata':False,'allow_partial':True,'content':'answer'}
        plan=api.plan_chat_export(data)
        assert plan['ok'] and plan['plan']['partial']
        result=api.confirm_chat_export(plan['plan']['plan_id'],True)['result']
        assert result['status']=='saved'
        raw=Path(result['path']).read_text(encoding='utf8')
        assert '未完成快照' in raw and '第一部分' in raw and '后续部分' not in raw
    finally:
        release.set(); finish(app)
    assert '后续部分' in api.conversation_snapshot()['result']['sessions'][0]['turns'][0]['answer']
    assert Path(result['path']).read_text(encoding='utf8')==raw


def test_export_reparse_directory_is_rejected_before_normalization(tmp_path,monkeypatch):
    app,api,_=setup(tmp_path)
    monkeypatch.setattr(api._connections,'build_model',lambda route:SimpleNamespace(invoke=lambda messages:SimpleNamespace(content='answer')))
    assert api.ask('q',{'stream':False})['ok']; finish(app)
    row=api.conversation_snapshot()['result']['sessions'][0]
    alias,actual=tmp_path/'alias',tmp_path/'actual'
    alias.mkdir(); actual.mkdir()
    app.window=SimpleNamespace(create_file_dialog=lambda *args:[str(alias)])
    resolve,lstat=Path.resolve,Path.lstat
    def linked_info(path,*args,**kwargs):
        info=lstat(path,*args,**kwargs)
        if path==alias:
            return SimpleNamespace(st_mode=info.st_mode,st_dev=info.st_dev,st_ino=info.st_ino,st_file_attributes=0x400)
        return info
    monkeypatch.setattr(Path,'resolve',lambda path,*args,**kwargs:actual if path==alias else resolve(path,*args,**kwargs))
    monkeypatch.setattr(Path,'lstat',linked_info)
    result=api.plan_chat_export({'session_id':row['id'],'turn_ids':[row['turns'][0]['id']],'format':'txt','metadata':False})
    assert not result['ok'] and result['reason']=='evidence_path_changed'
    assert list(actual.iterdir())==[]


def test_local_limit_redacts_before_cutting_a_credential(tmp_path,monkeypatch):
    app,api,_=setup(tmp_path)
    secret=connection()['api_key']
    monkeypatch.setattr(api._connections,'build_model',lambda route:SimpleNamespace(invoke=lambda messages:SimpleNamespace(content='x'*31995+secret,response_metadata={'finish_reason':'stop'})))
    assert api.ask('safe',{'stream':False})['ok']; finish(app)
    answer=api.conversation_snapshot()['result']['sessions'][0]['turns'][0]['answer']
    assert secret[:5] not in answer and len(answer)<=32000
