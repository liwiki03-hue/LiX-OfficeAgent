"""Synthetic transport/model tests: never real providers, model downloads or weights."""
import json

import httpx2
import pytest

from office_agent.embeddings import EmbeddingConfig, EmbeddingError, Embedder


def config(**overrides):
    return EmbeddingConfig(**{'channel':'api','model':'text-embedding-v3','dimension':64,
                              'api_key':'fixture-private-key','base_url':'https://dashscope.aliyuncs.com/compatible-mode/v1',**overrides})


def row(index=0, vector=None):
    return {'index':index,'embedding':vector if vector is not None else [1]+[0]*63,'object':'embedding'}


def response(data, status=200, headers=None):
    body = data if type(data) is bytes else json.dumps(data).encode()
    return httpx2.Response(status,headers=headers,stream=httpx2.ByteStream(body))


def embedder(handler, **overrides):
    return Embedder(config(**overrides),transport_factory=lambda:httpx2.MockTransport(handler))


def test_api_batches_ten_preserves_response_indices_and_identity():
    calls = []
    def handler(request):
        payload = json.loads(request.content)
        calls.append(payload)
        assert request.url.path == '/compatible-mode/v1/embeddings'
        assert request.headers['authorization'] == 'Bearer fixture-private-key'
        assert request.headers['accept-encoding'] == 'identity'
        assert payload['model'] == 'text-embedding-v3' and payload['dimensions'] == 64
        return response({'model':payload['model'],'data':[row(i) for i in reversed(range(len(payload['input'])))]})
    provider = embedder(handler)
    vectors = provider.embed([f'正文{i}' for i in range(23)])
    assert [len(call['input']) for call in calls] == [10,10,3]
    assert len(vectors) == 23 and all(len(vector) == 64 for vector in vectors)
    assert provider.identity == {'channel':'api','model':'text-embedding-v3','dimension':64}
    assert 'fixture-private-key' not in repr(provider.config)


@pytest.mark.parametrize('data',[{'data':[]},{'data':[row(vector=[1,0])]},
    {'data':[row(vector=[0]*64)]},{'data':[row(vector=[float('nan')]+[0]*63)]},
    {'data':[row(vector=[True]+[0]*63)]},{'data':[row(1)]},
    {'data':[row(0),row(0)]},{'model':'changed-model','data':[row()]},
    b'{"data":[],"data":[]}'])
def test_malformed_provider_response_is_redacted(data):
    with pytest.raises(EmbeddingError) as failure:
        embedder(lambda request:response(data)).embed(['正文'])
    assert 'fixture-private-key' not in str(failure.value)


@pytest.mark.parametrize('status,headers,body',[(403,{},b'fixture-private-provider-error'),
    (302,{'location':'https://fixture.invalid'},b''),
    (200,{'content-encoding':'gzip'},b'compressed'),
    (200,{'content-length':'2000001'},b''),(200,{},b'x'*2_000_001)],
    ids=['http-error','redirect','compressed','declared-large','stream-large'])
def test_redirects_compression_errors_and_large_bodies_fail_closed(status,headers,body):
    calls = []
    def handler(request):
        calls.append(request)
        return response(body,status,headers)
    with pytest.raises(EmbeddingError) as failure:
        embedder(handler).embed(['正文'])
    assert len(calls) == 1
    assert 'fixture-private-provider-error' not in str(failure.value)


@pytest.mark.parametrize('texts',[[],[''],['x\x00y'],['x'*8001],['API_KEY=fixture-private-key'],['sk-privatefixture123456'],['text']*501])
def test_invalid_or_sensitive_inputs_never_create_transport(texts):
    calls = []
    provider = Embedder(config(),transport_factory=lambda:calls.append('created'))
    with pytest.raises(EmbeddingError): provider.embed(texts)
    assert calls == []


@pytest.mark.parametrize('overrides',[{'base_url':'http://dashscope.aliyuncs.com/compatible-mode/v1'},
    {'base_url':'https://user:password@dashscope.aliyuncs.com/compatible-mode/v1'},
    {'base_url':'https://dashscope.aliyuncs.com/compatible-mode/v1?token=private'},
    {'base_url':'https://other.invalid/compatible-mode/v1'}, {'dimension':2},{'model':'unselected-model'},
    {'api_key':''},{'timeout':21},{'budget':61}])
def test_configuration_comes_from_canonical_values_and_rejects_bad_api(overrides):
    with pytest.raises(EmbeddingError): config(**overrides)


def test_from_values_never_uses_ambient_key(monkeypatch):
    monkeypatch.setenv('DASHSCOPE_API_KEY','fixture-ambient-private-key')
    values = {'OFFICE_EMBEDDING_CHANNEL':'api','OFFICE_EMBEDDING_MODEL':'text-embedding-v3',
              'OFFICE_EMBEDDING_DIMENSION':'64','DASHSCOPE_BASE_URL':'https://dashscope.aliyuncs.com/compatible-mode/v1'}
    with pytest.raises(EmbeddingError): EmbeddingConfig.from_values(values)
    values['DASHSCOPE_API_KEY'] = 'fixture-private-key'
    assert EmbeddingConfig.from_values(values).dimension == 64


def test_total_budget_checked_after_response_headers(monkeypatch):
    clock = [0.0]
    monkeypatch.setattr('office_agent.embeddings.time.monotonic',lambda:clock[0])
    def handler(request):
        clock[0] = 31
        return response({'data':[row()]})
    with pytest.raises(EmbeddingError,match='embedding_timeout'):
        embedder(handler).embed(['正文'])


def local(tmp_path, **overrides):
    root = tmp_path / 'offline-bge-fixture'
    root.mkdir()
    (root / 'model.safetensors').write_bytes(b'synthetic-not-real-model')
    return root, EmbeddingConfig('local',str(root),2,**overrides)


def test_local_factory_offline_safetensors_flags_lazy_load_and_pinned_identity(tmp_path):
    root,cfg = local(tmp_path)
    calls = []
    class Model:
        def encode(self,texts,**kwargs):
            calls.append(('encode',texts,kwargs))
            return [[1,0] for _ in texts]
    def factory(*args,**kwargs):
        calls.append(('load',args,kwargs))
        return Model()
    provider = Embedder(cfg,local_factory=factory)
    identity = provider.identity
    assert calls == [] and identity['model'].startswith('local:sha256:')
    assert provider.embed(['正文']*12) == [[1,0]]*12
    assert calls[0][2]['local_files_only'] is True
    assert calls[0][2]['trust_remote_code'] is False
    assert calls[0][2]['model_kwargs']['use_safetensors'] is True
    assert [len(call[1]) for call in calls if call[0] == 'encode'] == [10,2]
    (root / 'model.safetensors').write_bytes(b'changed-synthetic-model')
    with pytest.raises(EmbeddingError,match='embedding_local_changed'): provider.embed(['正文'])
    assert len([call for call in calls if call[0] == 'load']) == 1


@pytest.mark.parametrize('filename,content',[('weights.bin',b'pickle-like'),('custom.py',b'code'),
    ('modules.json',b'[{"type":"untrusted.Custom","path":""}]'),
    ('modules.json',b'[{"type":"sentence_transformers.models.Transformer","path":"../outside"}]'),
    ('tokenizer_config.json',b'{"tokenizer_file":"../outside.json"}')])
def test_local_bundle_rejects_code_pickle_modules_and_external_paths_before_factory(tmp_path,filename,content):
    root,cfg = local(tmp_path)
    (root / filename).write_bytes(content)
    calls = []
    provider = Embedder(cfg,local_factory=lambda *args,**kwargs:calls.append('load'))
    with pytest.raises(EmbeddingError): provider.embed(['正文'])
    assert calls == []


def test_missing_local_directory_is_redacted(tmp_path):
    provider = Embedder(EmbeddingConfig('local',str(tmp_path / 'private-missing-directory'),2))
    with pytest.raises(EmbeddingError) as failure: _ = provider.identity
    assert 'private-missing-directory' not in str(failure.value)
