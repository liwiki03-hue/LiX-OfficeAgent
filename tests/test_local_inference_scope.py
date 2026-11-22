import json
from pathlib import Path
from types import SimpleNamespace

import httpx
import pytest

from office_agent.webui import model_connections as connections


def route(**changes):
    return SimpleNamespace(**{'protocol': 'ollama', 'model_id': 'sample:7b',
                              'base_url': 'http://127.0.0.1:11434', 'api_key': '', **changes})


def confirm(candidate):
    return connections.ModelConnections.confirmed_local_inference(None, candidate)


@pytest.mark.parametrize('changes', [
    {'protocol':'openai'}, {'base_url':'https://ollama.com'}, {'base_url':'http://127.0.0.1:11434/v1'},
    {'base_url':'http://127.0.0.1:9999'}, {'api_key':'fixture'}, {'model_id':'large:cloud'},
])
def test_remote_or_relay_routes_never_even_request_local_metadata(changes, monkeypatch):
    calls = []
    monkeypatch.setattr(connections, '_ollama_listener', lambda: calls.append(True))
    assert not confirm(route(**changes)) and not calls


def test_identity_and_gguf_are_required_each_time(tmp_path, monkeypatch):
    root = tmp_path / 'models'
    blobs = root / 'blobs'
    blobs.mkdir(parents=True)
    blob = blobs / ('sha256-' + 'a' * 64)
    blob.write_bytes(b'GGUF' + bytes(20))
    monkeypatch.setenv('OLLAMA_MODELS', str(root))
    meta = {'details': {'format':'gguf'}, 'model_info': {'general.architecture':'qwen2', 'general.parameter_count':12},
            'modelfile': 'FROM ' + str(blob) + '\n'}
    record, listeners = [], [43]
    monkeypatch.setattr(connections, '_ollama_listener', lambda: listeners.pop(0) if len(listeners)>1 else listeners[0])
    def respond(request):
        record.append((request.method, str(request.url), json.loads(request.content), 'authorization' in request.headers))
        return httpx.Response(200, stream=httpx.ByteStream(json.dumps(meta).encode()), headers={'content-type':'application/json'})
    client = httpx.AsyncClient
    def direct(**kwargs):
        assert kwargs['trust_env'] is False and kwargs['follow_redirects'] is False
        return client(transport=httpx.MockTransport(respond), **kwargs)
    monkeypatch.setattr(httpx, 'AsyncClient', direct)
    assert confirm(route())
    assert record == [('POST','http://127.0.0.1:11434/api/show',{'model':'sample:7b'},False)]
    listeners[:] = [43,55]
    assert not confirm(route())
    listeners[:] = [None]
    count = len(record)
    assert not confirm(route()) and len(record) == count
    listeners[:] = [43]
    meta['remote_host'] = 'https://ollama.com'
    assert not confirm(route())
    del meta['remote_host']
    blob.write_bytes(b'other')
    assert not confirm(route())
    meta['modelfile'] = 'FROM ' + str(Path(tmp_path) / 'unknown')
    assert not confirm(route())


def test_probe_rejects_redirects_and_limits_reply_size(monkeypatch):
    monkeypatch.setattr(connections, '_ollama_listener', lambda: 43)
    responses = [httpx.Response(302, headers={'location':'https://example.invalid/'}),
                 httpx.Response(200, stream=httpx.ByteStream(b' ' * (256*1024+1)), headers={'content-type':'application/json'})]
    client = httpx.AsyncClient
    called = []
    def respond(request):
        called.append(request.url.host)
        return responses.pop(0)
    monkeypatch.setattr(httpx, 'AsyncClient', lambda **kwargs: client(transport=httpx.MockTransport(respond), **kwargs))
    assert not confirm(route()) and not confirm(route())
    assert called == ['127.0.0.1','127.0.0.1']
