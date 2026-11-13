from copy import deepcopy
import json
import threading
from types import SimpleNamespace

import httpx2
import pytest

from office_agent.literature import LiteratureError, LiteratureQuery, doi, normalize, group_records, csl
from office_agent.literature_client import LiteratureClient, LiteratureConfig
from office_agent.webui.literature import Literature
from office_agent.writes import Editor
from test_embeddings import response as raw_response
from test_fidelity import pdf_bytes


def response(data, status=200, headers=None):
    return raw_response(data, status, headers or {'content-type': 'application/json'})


def crossref(title='Synthetic thermal experiment', doi_value='10.1234/EXPT.01'):
    return {'DOI': doi_value, 'title': [title], 'author': [{'given': 'Ada', 'family': 'Lovelace'}],
            'type': 'journal-article', 'issued': {'date-parts': [[2024]]},
            'abstract': '<jats:p>Measurements at 20 ms.</jats:p>', 'reference': [{'DOI': '10.1234/related'}],
            'relation': {'is-preprint-of': [{'id': '10.1234/final'}]}}


def openalex(title='Synthetic thermal experiment'):
    return {'id': 'https://openalex.org/W12345', 'doi': 'https://doi.org/10.1234/expt.01', 'display_name': title,
            'publication_year': 2025, 'type': 'article', 'is_retracted': True,
            'authorships': [{'author': {'display_name': 'Ada Lovelace'}}],
            'abstract_inverted_index': {'Measurements': [0], 'at': [1], '20': [2], 'ms.': [3]},
            'best_oa_location': {'is_oa': True, 'pdf_url': 'https://arxiv.org/pdf/2401.12345', 'license': 'cc-by'},
            'referenced_works': ['https://openalex.org/W12346']}


def client_fixture(handler=None, **kwargs):
    requests = []
    now = [0.0]
    def default(request):
        requests.append(request)
        if handler:
            return handler(request)
        if request.url.host == 'api.crossref.org':
            return response({'message': {'items': [crossref()], 'total-results': 21}})
        return response({'results': [openalex()], 'meta': {'count': 1}})
    client = LiteratureClient(LiteratureConfig(**kwargs), transport_factory=lambda: httpx2.MockTransport(default), clock=lambda: now[0])
    return client, requests, now


def facade(tmp_path, client=None):
    editor = Editor(tmp_path, require_confirmation=True)
    selected = tmp_path / 'selected.pdf'
    selected.write_bytes(pdf_bytes(text='Synthetic evidence 20 ms.'))
    authenticated = [True]
    scope = lambda: (authenticated[0], False, (selected, editor.space.root))
    service = Literature(threading.RLock(), scope, lambda: editor, tmp_path)
    client, requests, now = client_fixture() if client is None else client
    service.factory = lambda: client
    return service, authenticated, requests, now


def search(service):
    reply = service.search({'query': 'thermal experiment', 'sources': ['crossref', 'openalex']})
    assert reply['ok'], reply
    return reply['result']


def test_doi_normalization_conflicts_and_relations_remain_separate():
    assert doi('https://doi.org/10.1234/EXPT.01') == '10.1234/expt.01'
    first, second = normalize(crossref(), 'crossref'), normalize(openalex(), 'openalex')
    groups = group_records([first, second])
    assert len(groups) == 1 and len(groups[0]['variants']) == 2
    assert 'year' in groups[0]['conflicts'] and 'kind' in groups[0]['conflicts']
    assert first['relations'][0]['type'] == 'is-preprint-of'
    assert first['abstract'] == 'Measurements at 20 ms.'
    assert second['retraction_status'] == 'reported_retracted' and not first['support_assessed']
    assert first['fulltext_state'] == 'not_acquired' and second['access'] == 'abstract'
    assert csl(first)['author'] == [{'family': 'Lovelace', 'given': 'Ada'}]
    assert csl(second)['author'] == [{'literal': 'Ada Lovelace'}]
    no_doi = normalize({**crossref(), 'DOI': None}, 'crossref')
    alternative = {**no_doi, 'id': 'lit:' + 'b' * 64}
    separate = group_records([no_doi, alternative])
    assert len(separate) == 2 and separate[0]['possible_duplicates'] == [alternative['id']]


@pytest.mark.parametrize('payload', [
    {'query': 'x', 'year_from': 2025, 'year_to': 2024}, {'query': 'x', 'page': 4},
    {'query': 'x', 'author': 'name,type:book'}, {'query': 'x', 'sources': ('other',)},
    {'query': '10.1234/test', 'mode': 'doi', 'year_from': 2024}, {'query': 'x', 'path': 'C:/outside'}
])
def test_invalid_query_does_not_reach_transport(payload):
    with pytest.raises((ValueError, LiteratureError)):
        LiteratureQuery.model_validate(payload)


def test_fixed_services_filters_cache_and_budget_do_not_send_keys_by_default():
    client, calls, now = client_fixture(openalex_key='fixture-never-send')
    query = LiteratureQuery(query='thermal', author='Ada Lovelace', year_from=2020, year_to=2025, kind='article')
    rows, reports = client.search(query)
    assert len(rows) == 2 and len(calls) == 2
    assert all(request.url.scheme == 'https' and 'authorization' not in request.headers for request in calls)
    assert 'fixture-never-send' not in repr(client) + repr(client.config) + json.dumps(reports)
    assert calls[0].url.params['filter'] == 'from-pub-date:2020-01-01,until-pub-date:2025-12-31,type:journal-article'
    assert 'raw_author_name.search:Ada Lovelace' in calls[1].url.params['filter']
    cached, statuses = client.search(query)
    assert len(calls) == 2 and all(row['cached'] for row in statuses)
    cached[0]['title'] = 'tampered'
    assert client.search(query)[0][0]['title'] != 'tampered'
    client.calls = client.config.request_limit
    now[0] = 200
    assert all(row['status'] == 'literature_budget' for row in client.search(query)[1])


@pytest.mark.parametrize('status,code', [(401, 'literature_auth_required'), (429, 'literature_rate_limited'),
                                        (302, 'literature_redirect_refused'), (500, 'literature_service_failed')])
def test_known_service_failures_are_redacted_and_not_retried(status, code):
    client, calls, now = client_fixture(lambda request: response({'private': 'must-not-escape'}, status=status, headers={'content-type':'application/json', 'location':'https://other.invalid', 'retry-after':'3'}))
    rows, reports = client.search(LiteratureQuery(query='thermal', sources=('crossref',)))
    assert rows == [] and reports[0]['status'] == code and len(calls) == 1
    assert 'must-not-escape' not in json.dumps(reports)


def test_timeout_malformed_response_and_key_request_quota():
    def timeout(request):
        raise httpx2.ReadTimeout('fixture-sensitive-body')
    client, calls, now = client_fixture(timeout)
    assert client.search(LiteratureQuery(query='thermal', sources=('crossref',)))[1][0]['status'] == 'literature_timeout'
    client, calls, now = client_fixture(openalex_key='fixture-explicit-key', key_requests=1)
    client.search(LiteratureQuery(query='thermal', sources=('openalex',)))
    assert calls[0].headers['authorization'] == 'Bearer fixture-explicit-key' and 'api_key' not in str(calls[0].url)
    now[0] = 2
    assert client.search(LiteratureQuery(query='different', sources=('openalex',)))[1][0]['status'] == 'literature_budget'
    assert len(calls) == 1


def test_session_review_export_cancel_duplicate_refusal_and_verified_rollback(tmp_path):
    service, authenticated, calls, now = facade(tmp_path)
    result = search(service)
    first, second = [row['id'] for row in result['groups'][0]['variants']]
    assert service.review(first, 'include', '')['reason'] == 'literature_invalid'
    assert service.review(first, 'include', 'Methods relevant')['ok']
    assert service.review(second, 'exclude', 'Different year')['ok']
    plan = service.export('csl')['result']['plan']
    assert not list(tmp_path.glob('.office-rag-*'))
    assert service.confirm(plan['plan_id'], False)['result']['wrote_nothing']
    plan = service.export('csl')['result']['plan']
    saved = service.confirm(plan['plan_id'], True)['result']
    target = tmp_path / saved['evidence']
    assert target.name == 'references.csl.json' and json.loads(target.read_bytes())[0]['DOI'] == '10.1234/expt.01'
    assert service.rollback(saved['save_id'], False)['result']['wrote_nothing']
    assert service.rollback(saved['save_id'], True)['result']['status'] == 'rolled_back'
    assert not target.exists() and (target.parent / 'audit-removed.json').exists()
    assert service.review(second, 'include', 'Compare sources')['ok']
    assert service.export('csl')['reason'] == 'literature_duplicate_selection'
    checked = service.check_citations('Claim [1], unrelated [8].')['result']
    assert checked['unknown'] == [8] and checked['duplicate_dois'] == ['10.1234/expt.01'] and not checked['support_assessed']
    plan = service.export('record')
    assert not plan['ok']
    service.invalidate()
    assert not service.records and not service.saves and not service.plans
    authenticated[0] = False
    assert service.search({'query': 'thermal'})['reason'] == 'not_authenticated'


def test_late_network_result_and_duplicate_click_do_not_cross_logout(tmp_path):
    started, release = threading.Event(), threading.Event()
    def blocking(request):
        started.set()
        assert release.wait(3)
        return response({'message': {'items': [crossref()], 'total-results': 1}})
    fixture = client_fixture(blocking)
    service, authenticated, calls, now = facade(tmp_path, fixture)
    replies = []
    worker = threading.Thread(target=lambda: replies.append(service.search({'query':'thermal', 'sources':['crossref']})))
    worker.start()
    try:
        assert started.wait(2)
        assert service.search({'query':'thermal'})['reason'] == 'busy'
        authenticated[0] = False
        service.invalidate()
    finally:
        release.set()
        worker.join(5)
    assert not worker.is_alive() and replies[0]['reason'] == 'literature_scope_changed'
    assert not service.records and not service.reports and not service.running


def test_public_fulltext_boundaries_save_parse_and_rollback(tmp_path):
    data = pdf_bytes(text='Synthetic public evidence 20 ms.')
    def handler(request):
        if request.url.host == 'arxiv.org':
            assert 'authorization' not in request.headers and not request.url.query
            return httpx2.Response(200, headers={'content-type':'application/pdf'}, stream=httpx2.ByteStream(data))
        return response({'results': [openalex()], 'meta': {'count': 1}})
    fixture = client_fixture(handler)
    service, authenticated, calls, now = facade(tmp_path, fixture)
    result = service.search({'query':'thermal', 'sources':['openalex']})['result']
    record_id = result['groups'][0]['variants'][0]['id']
    plan = service.download(record_id)['result']['plan']
    assert plan['filename'] == 'fulltext.pdf' and not list(tmp_path.glob('.office-rag-*'))
    saved = service.confirm(plan['plan_id'], True)['result']
    target = tmp_path / saved['evidence']
    assert target.read_bytes() == data
    before_repeat = len(calls)
    assert service.download(record_id)['reason'] == 'literature_fulltext_exists' and len(calls) == before_repeat
    audit = json.loads((target.parent / 'audit-verified.json').read_bytes())
    assert audit['operation'] == 'save_literature' and audit['provenance']['source_url'] == 'https://arxiv.org/pdf/2401.12345'
    assert audit['provenance']['license'] == 'cc-by' and audit['provenance']['record_id'] == record_id
    parsed = service.read_fulltext(saved['save_id'])['result']
    assert 'Synthetic public evidence' in parsed['pages'][0]['text'] and not parsed['support_assessed']
    assert parsed['source_document']['literature_id'] == record_id and parsed['source_document']['access'] == 'public_fulltext'
    assert parsed['source_document']['source_sha256'] == parsed['source_sha256']
    service._accept([normalize(openalex(), 'openalex')], [])
    assert service.records[record_id]['fulltext_state'] == 'parsed_text_layer'
    assert service.records[record_id]['fulltext_source']['source_sha256'] == parsed['source_sha256']
    assert service.rollback(saved['save_id'], True)['result']['status'] == 'rolled_back' and not target.exists()
    assert service.records[record_id]['fulltext_state'] == 'not_acquired'
    row = service.records[record_id]
    row['public_fulltext']['url'] = 'https://127.0.0.1/private.pdf'
    now[0] = 2
    before = len(calls)
    assert service.download(record_id)['reason'] == 'literature_fulltext_refused' and len(calls) == before
    row['public_fulltext']['url'] = 'https://arxiv.org/pdf/2401.12345'
    row['public_fulltext']['license'] = None
    assert service.download(record_id)['reason'] == 'literature_fulltext_refused'


def test_local_pdf_association_does_not_claim_doi_match_or_write(tmp_path):
    service, authenticated, calls, now = facade(tmp_path)
    result = search(service)
    record_id = result['groups'][0]['variants'][0]['id']
    local = service.local_fulltext(record_id)['result']
    assert local['blocks'] and not local['doi_match_verified'] and local['access'] == 'user_local'
    assert local['source_document']['literature_id'] == record_id and local['source_document']['access'] == 'user_local'
    assert len(calls) == 2 and not list(tmp_path.glob('.office-rag-*'))


def test_metadata_revision_cannot_silently_change_an_included_record(tmp_path):
    service, authenticated, calls, now = facade(tmp_path)
    first = normalize(crossref('Original title'), 'crossref')
    changed = normalize(crossref('Changed title'), 'crossref')
    assert first['id'] != changed['id'] and first['source_record_sha256'] != changed['source_record_sha256']
    search(service)
    service.records = {first['id']: first}
    assert service.review(first['id'], 'include', 'Reviewed this version')['ok']
    service._accept([changed], [])
    plan = service.export('csl')['result']['plan']
    outcome = service.confirm(plan['plan_id'], True)['result']
    assert json.loads((tmp_path / outcome['evidence']).read_bytes())[0]['title'] == 'Original title'


@pytest.mark.parametrize('damage', ['expiry', 'tamper'])
def test_export_expiry_and_tamper_reject_without_deleting_output(tmp_path, damage):
    service, authenticated, calls, now = facade(tmp_path)
    result = search(service)
    record_id = result['groups'][0]['variants'][0]['id']
    service.review(record_id, 'include', 'Reviewed')
    plan = service.export('csl')['result']['plan']
    archive = service.archives['csl']
    if damage == 'expiry':
        archive._plans[plan['plan_id']]['time'] -= 301
        reply = service.confirm(plan['plan_id'], True)['result']
        assert reply['status'] == 'rejected' and reply['wrote_nothing']
        assert not list(tmp_path.glob('.office-rag-*'))
    else:
        saved = service.confirm(plan['plan_id'], True)['result']
        target = tmp_path / saved['evidence']
        target.write_bytes(b'External modification')
        reply = service.rollback(saved['save_id'], True)
        assert reply['reason'] == 'evidence_path_changed' and target.read_bytes() == b'External modification'


@pytest.mark.parametrize('bad', [b'{broken', b'{"message":{},"message":{}}'])
def test_malformed_metadata_keeps_stable_redacted_error(bad):
    client, calls, now = client_fixture(lambda request: response(bad))
    rows, reports = client.search(LiteratureQuery(query='thermal', sources=('crossref',)))
    assert rows == [] and reports[0]['status'] == 'literature_response' and len(calls) == 1


def test_failed_search_record_export_preserves_error_and_expression(tmp_path):
    fixture = client_fixture(lambda request: response({}, status=503))
    service, authenticated, calls, now = facade(tmp_path, fixture)
    search(service)
    plan = service.export('record')['result']['plan']
    saved = service.confirm(plan['plan_id'], True)['result']
    exported = json.loads((tmp_path / saved['evidence']).read_bytes())
    assert exported['searches'][0]['expression']['query'] == 'thermal experiment'
    assert exported['searches'][0]['status'] == 'literature_service_failed' and not exported['systematic_review']


def test_post_admission_write_failure_fuses_service(tmp_path, monkeypatch):
    service, authenticated, calls, now = facade(tmp_path)
    result = search(service)
    record_id = result['groups'][0]['variants'][0]['id']
    service.review(record_id, 'include', 'Reviewed')
    plan = service.export('csl')['result']['plan']
    monkeypatch.setattr(service.archives['csl'], '_save_audit', lambda plan: (_ for _ in ()).throw(RuntimeError('private-message')))
    result = service.confirm(plan['plan_id'], True)['result']
    assert result['status'] == 'uncertain' and result['applied'] is None and not result['wrote_nothing']
    assert service.disabled and service.export('csl')['reason'] == 'literature_uncertain'
    assert 'private-message' not in json.dumps(result)


@pytest.mark.parametrize('kind', ['redirect', 'oversized', 'wrong_mime'])
def test_pdf_fetch_rejects_redirect_size_and_mime_without_following(kind):
    def handler(request):
        if kind == 'redirect':
            return httpx2.Response(302, headers={'location':'https://private.invalid/file.pdf'})
        return httpx2.Response(200, headers={'content-type':'text/html' if kind == 'wrong_mime' else 'application/pdf',
                                           'content-length':str(512 * 1024 + 1 if kind == 'oversized' else 5)},
                               stream=httpx2.ByteStream(b'%PDF-'))
    client, calls, now = client_fixture(handler)
    with pytest.raises(LiteratureError, match={'redirect':'literature_redirect_refused', 'oversized':'literature_capacity', 'wrong_mime':'literature_response'}[kind]):
        client.fulltext(normalize(openalex(), 'openalex'))
    assert len(calls) == 1 and calls[0].url.host == 'arxiv.org'


def test_replaced_workspace_directory_rejects_before_network_or_write(tmp_path):
    root = tmp_path / 'granted'
    root.mkdir()
    service, authenticated, calls, now = facade(root)
    search(service)
    root.rename(tmp_path / 'retired')
    root.mkdir()
    (root / 'selected.pdf').write_bytes(pdf_bytes())
    before = len(calls)
    assert service.search({'query':'different'})['reason'] == 'literature_scope_changed'
    assert len(calls) == before and not list(root.glob('.office-rag-*'))


def test_bibliographic_fields_survive_csl_export_without_guessed_author_parts():
    record = normalize({**crossref(), 'container-title':['Technical Experiments'], 'volume':'12',
                        'issue':'3', 'page':'20-24', 'publisher':'Synthetic Publisher'}, 'crossref')
    value = csl(record)
    assert value['container-title'] == 'Technical Experiments' and value['volume'] == '12'
    assert value['issue'] == '3' and value['page'] == '20-24' and value['publisher'] == 'Synthetic Publisher'
    assert 'literal' in csl(normalize(openalex(), 'openalex'))['author'][0]


def test_untrusted_page_field_cannot_bypass_text_redaction():
    record = {**openalex(), 'biblio': {'first_page':'api_key=fixture-sensitive-value'}}
    with pytest.raises(LiteratureError, match='literature_sensitive'):
        normalize(record, 'openalex')


def test_downloaded_empty_pdf_does_not_claim_body_evidence(tmp_path):
    service, authenticated, calls, now = facade(tmp_path)
    result = search(service)
    record = next(row for row in result['groups'][0]['variants'] if row['source'] == 'openalex')
    service.client.fulltext = lambda record, **kwargs: pdf_bytes(text='')
    plan = service.download(record['id'])['result']['plan']
    saved = service.confirm(plan['plan_id'], True)['result']
    parsed = service.read_fulltext(saved['save_id'])['result']
    assert not parsed['text_layer_available'] and not parsed['support_assessed']
    assert service.records[record['id']]['fulltext_state'] == 'text_layer_empty_in_preview'


def test_citing_expansion_is_single_bounded_metadata_request():
    client, calls, now = client_fixture()
    rows, report = client.citing(normalize(openalex(), 'openalex'))
    assert len(calls) == 1 and calls[0].url.params['filter'] == 'cites:W12345'
    assert len(rows) == 1 and report['returned'] == 1 and not rows[0]['support_assessed']
    assert report['auth_mode'] == 'anonymous'
    keyed, keyed_calls, now = client_fixture(openalex_key='fixture-key-mode', key_requests=1)
    assert keyed.citing(normalize(openalex(), 'openalex'))[1]['auth_mode'] == 'configured_key'
    assert keyed_calls[0].headers['authorization'] == 'Bearer fixture-key-mode'
