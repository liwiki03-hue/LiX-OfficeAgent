import pytest

from office_agent.webui.skill_discovery import DiscoveryError, SkillDiscovery, normalize


def response():
    return {'skills': [{'id': 'example/workflows/document-review', 'source': 'example/workflows',
                        'skillId': 'document-review', 'name': 'Document review'}]}


def test_search_cache_does_not_turn_failure_into_empty_success():
    calls = []
    def fetch(query):
        calls.append(query)
        if len(calls) > 1:
            raise DiscoveryError('rate_limited')
        return response()
    service = SkillDiscovery(fetch)
    first = service.search('document')
    assert first['ok'] and len(first['items']) == 1
    assert first['items'][0]['commit'] is None
    assert first['items'][0]['compatibility'] == 'not_checked'
    first['items'].clear()
    cached = service.search('document')
    assert not cached['ok'] and cached['cached'] and len(cached['items']) == 1
    assert cached['reason'] == 'rate_limited'
    missing = service.search('other')
    assert not missing['ok'] and missing['items'] == []


def test_invalid_query_and_candidate_are_not_download_authority():
    calls = []
    service = SkillDiscovery(lambda query: calls.append(query))
    for query in ['', 'x', 'x' * 121, 'abc\nfile', None]:
        assert service.search(query)['reason'] == 'invalid_query'
    assert calls == []
    for source in ['../workflows', 'https://example.com', 'example/workflows/more', 'example/..']:
        payload = response()
        payload['skills'][0]['source'] = source
        with pytest.raises(DiscoveryError):
            normalize(payload)


def test_results_are_bounded_and_empty_is_distinct():
    assert SkillDiscovery(lambda query: {'skills': []}).search('document')['ok']
    with pytest.raises(DiscoveryError):
        normalize({'error': 'not a directory'})
    rows = [{'id': f'example/workflows/skill-{i}', 'source': 'example/workflows',
             'skillId': f'skill-{i}', 'name': f'Skill {i}'} for i in range(50)]
    assert len(normalize({'skills': rows})) == 24


def test_busy_search_does_not_queue_requests():
    service = SkillDiscovery(lambda query: response())
    service._slots.acquire()
    service._slots.acquire()
    try:
        assert service.search('document')['reason'] == 'busy'
    finally:
        service._slots.release()
        service._slots.release()
