import json

import httpx2
import pytest

from office_agent.research_model import ResearchModel, ResearchModelConfig, ResearchModelError


def config(**values):
    return ResearchModelConfig('ollama', 'fixture-model', 'http://127.0.0.1:11434', 8192, **values)


def test_explicit_bounded_model_request_no_tools_and_cancellation():
    calls = []

    def handle(request):
        calls.append(json.loads(request.content))
        body = json.dumps({'message': {'content': json.dumps({'cards': [], 'suggestions': [], 'missing_evidence': [], 'stop_reason': 'insufficient'})}}).encode()
        return httpx2.Response(200, stream=httpx2.ByteStream(body))

    model = ResearchModel(config(), httpx2.MockTransport(handle))
    assert not calls
    assert model.suggest({'question': 'Technical fixture'})['stop_reason'] == 'insufficient'
    assert len(calls) == 1 and not calls[0].get('tools') and calls[0]['options']['num_predict'] == 1800
    with pytest.raises(ResearchModelError, match='research_cancelled'):
        model.suggest({'question': 'fixture'}, lambda: True)
    assert len(calls) == 1


@pytest.mark.parametrize('status', [301, 401, 500])
def test_http_errors_and_redirects_do_not_echo_response_or_retry(status):
    seen = []

    def handle(request):
        seen.append(request)
        return httpx2.Response(status, headers={'Location': 'https://other.example/'}, content=b'private-provider-response')

    with pytest.raises(ResearchModelError, match='^research_model_service$'):
        ResearchModel(config(), httpx2.MockTransport(handle)).suggest({'question': 'fixture'})
    assert len(seen) == 1


@pytest.mark.parametrize('address', ['http://example.org', 'https://user:pass@example.org', 'https://example.org?q=secret', 'https://example.org/#token', 'https://example.org:0', 'https://example.org/a b', 'https://example.org\\a'])
def test_model_url_boundary_matches_settings(address):
    with pytest.raises(ResearchModelError, match='research_model_configuration'):
        ResearchModelConfig('ollama', 'fixture-model', address, 8192)


def test_model_timeout_response_and_context_are_sanitized():
    def timeout(request):
        raise httpx2.ReadTimeout('private-request-header-details')

    with pytest.raises(ResearchModelError, match='^research_model_timeout$'):
        ResearchModel(config(), httpx2.MockTransport(timeout)).suggest({'question': 'fixture'})
    with pytest.raises(ResearchModelError, match='research_model_context_budget'):
        ResearchModel(ResearchModelConfig('ollama', 'fixture-model', 'http://127.0.0.1:11434', 1024)).suggest({'question': 'fixture'})
