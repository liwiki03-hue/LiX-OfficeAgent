"""Synthetic canonical settings fixtures only: no live secrets/models/Office."""
import json
import os

import pytest
from office_agent import settings_store as m
from office_agent.embeddings import EmbeddingConfig, EmbeddingError

NUMBERS = m.EMBEDDING_NUMBERS
BASE = b'# retained\r\nMODEL_PROVIDER=deepseek\r\nDEEPSEEK_MODEL=fixture\r\nUNRELATED=value\r\n'


def make(tmp_path, suffix=b''):
    (tmp_path / '.env').write_bytes(BASE + suffix)
    return m.SettingsStore(tmp_path)


def state(store):
    return {str(p.relative_to(store.root)): p.read_bytes() for p in store.root.rglob('*') if p.is_file()}


@pytest.mark.parametrize('values', [(1, 1, 1), (4096, 20, 60), (64, 1.25, 30.5), (128, 1.0000000000000002, 60.0)])
def test_typed_roundtrip_and_owned_rollback(tmp_path, values):
    store = make(tmp_path)
    changes = dict(zip(NUMBERS, values))
    plan = store.plan_update(changes, 'owner')
    assert (tmp_path / '.env').read_bytes() == BASE
    assert len(list(tmp_path.iterdir())) == 1
    assert plan['settings']['embedding_dimension'] == values[0]
    saved = store.confirm_update(plan['plan_id'], 'owner', True)
    assert saved['status'] == 'saved'
    actual = store.snapshot()
    assert type(actual['embedding_dimension']) is int
    for name, value in changes.items():
        assert actual[name] == value
        assert f'{NUMBERS[name][0]}={value}\r\n'.encode() in (tmp_path / '.env').read_bytes()
    rollback = store.plan_rollback(saved['save_id'], 'owner')
    assert store.confirm_rollback(rollback['rollback_id'], 'owner', True)['status'] == 'rolled_back'
    assert (tmp_path / '.env').read_bytes() == BASE
    assert store.confirm_rollback(rollback['rollback_id'], 'owner', True)['wrote_nothing'] is True


INVALID = [(name, value) for name, (_, upper, _) in NUMBERS.items()
           for value in (True, False, None, '1', 0, -1, upper + 1, float('nan'), float('inf'), float('-inf'), 10 ** 10000)]
INVALID += [('embedding_dimension', 1.0), ('embedding_dimension', 1.5), ('embedding_missing', 1)]


@pytest.mark.parametrize('name,value', INVALID, ids=[f'case-{i}' for i in range(len(INVALID))])
def test_bad_typed_parameters_no_write_no_grants(tmp_path, name, value):
    store = make(tmp_path)
    before = state(store)
    with pytest.raises(m.SettingsError, match='^settings_invalid$'):
        store.plan_update({name: value}, 'owner')
    assert state(store) == before and not store._pending and not store._issued


def test_missing_defaults_no_ambient_fallback_or_mutation(tmp_path, monkeypatch):
    for name, (env, _, _) in NUMBERS.items():
        monkeypatch.setenv(env, '999')
    monkeypatch.setenv('DASHSCOPE_API_KEY', 'ambient-not-a-real-key')
    before = dict(os.environ)
    store = make(tmp_path)
    snapshot = store.snapshot()
    assert snapshot['embedding_dimension'] is None
    assert snapshot['embedding_timeout'] == 20 and snapshot['embedding_budget'] == 30
    plan = store.plan_update({'model': 'another-fixture'}, 'owner')
    saved = store.confirm_update(plan['plan_id'], 'owner', True)
    rollback = store.plan_rollback(saved['save_id'], 'owner')
    assert store.confirm_rollback(rollback['rollback_id'], 'owner', True)['status'] == 'rolled_back'
    assert dict(os.environ) == before
    assert 'ambient-not-a-real-key' not in json.dumps([snapshot, plan])


RAW_INVALID = ['', '# empty', '""', '0', '-1', '1e0', '${OTHER}', '1suffix # injected\nOTHER=unexpected', '"1', '"1\n2"', '9' * 10000, '１', '1suffix']
BAD_RAW = [(name, value) for name in NUMBERS for value in RAW_INVALID]
BAD_RAW += [('embedding_dimension', '4097'), ('embedding_dimension', '1.0'), ('embedding_timeout', '20.1'), ('embedding_budget', '60.1')]


@pytest.mark.parametrize('name,raw', BAD_RAW, ids=[f'raw-{i}' for i in range(len(BAD_RAW))])
def test_malformed_current_config_fails_closed_no_write(tmp_path, name, raw):
    env = NUMBERS[name][0]
    store = make(tmp_path, f'{env}={raw}\n'.encode('utf-8'))
    before = state(store)
    for operation in (store.snapshot, lambda: store.plan_update({'model': 'another-fixture'}, 'owner')):
        with pytest.raises(m.SettingsError, match='^settings_file_invalid$'):
            operation()
    assert state(store) == before and not store._pending and not store._issued


@pytest.mark.parametrize('name', NUMBERS)
@pytest.mark.parametrize('form', ['bare', 'duplicate', 'bare-duplicate'])
def test_numeric_bare_or_duplicate_keys_rejected(tmp_path, name, form):
    env = NUMBERS[name][0]
    raw = env if form == 'bare' else f'{env}=1\nexport {env}="2"' if form == 'duplicate' else f'{env}\n{env}=1'
    store = make(tmp_path, (raw + '\n').encode())
    before = state(store)
    with pytest.raises(m.SettingsError, match='^settings_file_invalid$'):
        store.snapshot()
    assert state(store) == before


@pytest.mark.parametrize('name', NUMBERS)
@pytest.mark.parametrize('quote', ['', "'", '"'])
def test_quoted_or_plain_numbers_comments_preserved(tmp_path, name, quote):
    env = NUMBERS[name][0]
    store = make(tmp_path, f'{env}={quote}2{quote} # keep\r\n'.encode())
    assert store.snapshot()[name] == 2
    plan = store.plan_update({name: 3}, 'owner')
    result = store.confirm_update(plan['plan_id'], 'owner', True)
    assert result['status'] == 'saved'
    assert f'{env}=3 # keep\r\n'.encode() in (tmp_path / '.env').read_bytes()
    assert b'UNRELATED=value\r\n' in (tmp_path / '.env').read_bytes()


@pytest.mark.parametrize('approved', [False, 1, 'true', None, [], {}], ids=['false','one','string','null','list','dict'])
def test_numeric_literal_approvals_consumed_and_owned(tmp_path, approved):
    store = make(tmp_path)
    plan = store.plan_update({'embedding_dimension': 64}, 'owner')
    assert store.confirm_update(plan['plan_id'], 'owner', approved)['wrote_nothing'] is True
    assert store.confirm_update(plan['plan_id'], 'owner', True)['wrote_nothing'] is True
    assert (tmp_path / '.env').read_bytes() == BASE
    plan = store.plan_update({'embedding_dimension': 64}, 'owner')
    saved = store.confirm_update(plan['plan_id'], 'owner', True)
    with pytest.raises(m.SettingsError):
        store.plan_rollback(saved['save_id'], 'foreign')
    with pytest.raises(m.SettingsError):
        m.SettingsStore(tmp_path).plan_rollback(saved['save_id'], 'owner')
    rollback = store.plan_rollback(saved['save_id'], 'owner')
    current = state(store)
    assert store.confirm_rollback(rollback['rollback_id'], 'owner', approved)['wrote_nothing'] is True
    assert state(store) == current
    assert store.confirm_rollback(rollback['rollback_id'], 'owner', True)['wrote_nothing'] is True


@pytest.mark.parametrize('channel,model,dimension', [('local','C:/models/fixture',64), ('api','text-embedding-v4',64)])
def test_real_embedding_config_accepts_canonical_numbers_without_model_calls(tmp_path, channel, model, dimension):
    source = b'DASHSCOPE_API_KEY=synthetic-fixture-key\r\nDASHSCOPE_BASE_URL=https://dashscope.aliyuncs.com/compatible-mode/v1\r\n'
    store = make(tmp_path, source)
    plan = store.plan_update({'embedding_channel': channel, 'embedding_model': model, 'embedding_dimension': dimension, 'embedding_timeout': 1.25, 'embedding_budget': 30.5}, 'owner')
    assert 'synthetic-fixture-key' not in json.dumps([plan, store.snapshot()])
    assert store.confirm_update(plan['plan_id'], 'owner', True)['status'] == 'saved'
    config = EmbeddingConfig.from_values(store._values(store._read(tmp_path / '.env')))
    assert (config.channel, config.model, config.dimension, config.timeout, config.budget) == (channel, model, dimension, 1.25, 30.5)


def test_generic_settings_not_coupled_to_provider_specific_dimension(tmp_path):
    store = make(tmp_path)
    plan = store.plan_update({'embedding_channel': 'api', 'embedding_model': 'generic-fixture', 'embedding_dimension': 3}, 'owner')
    assert store.confirm_update(plan['plan_id'], 'owner', True)['status'] == 'saved'
    with pytest.raises(EmbeddingError):
        EmbeddingConfig.from_values(store._values(store._read(tmp_path / '.env')))
