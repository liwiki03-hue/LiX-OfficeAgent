import threading

import pytest

from office_agent.literature import normalize
from office_agent.skills_backend import SkillsBackend
from office_agent.webui.citation_turn import _present, _query_fields
from office_agent.webui.floating_session import FloatingApi
from office_agent.webui.literature import Literature
from office_agent.webui.office_workflows import workflow_package
from tests.test_document_conversation import Answers, attach, finish
from tests.test_literature import client_fixture, crossref, response
from tests.test_paged_editor import draft


def test_citation_fields_limit_outgoing_metadata_and_keep_conflicts():
    query, title, years, authors = _query_fields('核对选区', '作者：Grace Hopper；题名：《另一篇》；年份：2019；DOI: https://doi.org/10.1234/EXPT.01')
    assert (query.query, query.mode, query.sources) == ('10.1234/expt.01', 'doi', ('crossref',))
    row = normalize(crossref(), 'crossref')
    answer, matches = _present(query, title, years, authors, [row, row])
    assert matches == [row]
    assert all(word in answer for word in ['题名不符', '年份不符', '2019', '2024', '作者写法或顺序有差异', 'Grace Hopper', 'Ada Lovelace'])
    second = normalize(crossref(doi_value=None), 'crossref')
    query, title, years, authors = _query_fields('请核对《Synthetic thermal experiment》')
    answer, matches = _present(query, title, years, authors, [row, second])
    assert '多条候选' in answer and matches == [row, second] and 'DOI 缺失' in answer
    assert '没有查到' in _present(query, title, [], '', [])[0]
    for phrase in ['帮我核对这条文献', '核对《raw-contact@example.org》', '核对 10.1234/a 和 10.1234/b']:
        with pytest.raises(ValueError):
            _query_fields(phrase)


def test_word_citation_dispatch_once_and_followup_uses_only_metadata(draft):
    _, floating, source = draft
    model = Answers()
    attach(floating, model)
    original = source.read_bytes()
    client, requests, _ = client_fixture(lambda request: response({'message': crossref()}))
    editor = floating.app.editor
    service = Literature(threading.RLock(), lambda: (True, False, (source, editor.space.root)), lambda: editor, source.parent)
    service.factory = lambda: client
    floating.api._literature = service
    backend = SkillsBackend(source.parent)
    plan = backend.plan_public(workflow_package('citation-check'), 'session')
    assert backend.confirm_change(plan['plan_id'], 'session', True)['verified']
    choice = next(choice for choice in backend.turn_choices() if choice['id'] == 'user:lix-citation-check')
    ticket = {key: choice[key] for key in ('id', 'sha256')}
    chat = floating.conversation
    doc = floating.document

    def send(identifier, text, selected=True):
        result = FloatingApi.conversation_begin(floating.api, identifier, text, doc['id'], doc['version'], 0,
                                                local_only=True, resources={'skill': ticket} if selected else None)
        assert result['ok']
        finish(floating)
        return chat.turns[identifier]

    row = send('cancel-check', '核对 10.1234/EXPT.01')
    assert row['stage'] == 'scope_consent' and not requests
    token = chat.scopes.pending[row['id']].token
    assert chat.scopes.decide(token, 'cancel')['ok'] and not requests
    row = send('do-check', '核对 https://doi.org/10.1234/EXPT.01，年份 2019')
    token = chat.scopes.pending[row['id']].token
    assert '10.1234/expt.01' in row['text'] and '2019' not in chat.scopes.decide(token, 'preview')['preview']
    assert chat.scopes.decide(token, 'allow')['ok']
    with pytest.raises(ValueError):
        chat.scopes.decide(token, 'allow')
    finish(floating)
    assert row['status'] == 'complete' and len(requests) == 1 and '年份不符' in row['answer']
    assert row['citation_sources'][0]['url'] == 'https://doi.org/10.1234/expt.01' and row['citation_receipt']['response_sha256']
    reply = send('support-check', '它是否证明我的结论？', False)
    assert '仅取得了书目信息' in reply['answer'] and len(requests) == 1 and not model.calls
    assert source.read_bytes() == original and not floating.outputs
