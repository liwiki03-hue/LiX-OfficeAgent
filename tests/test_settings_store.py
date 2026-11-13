"""Isolated settings fixtures: never read project .env or invoke providers/Office."""
import json
import os
from pathlib import Path
import pytest
from office_agent import settings_store as m

OLD = b"# preserved\r\nMODEL_PROVIDER=deepseek\r\nDEEPSEEK_API_KEY=old-private-key\r\nDEEPSEEK_MODEL=old-model # keep comment\r\nUNKNOWN_SETTING=value\r\n"

@pytest.fixture
def store(tmp_path):
    (tmp_path / '.env').write_bytes(OLD)
    return m.SettingsStore(tmp_path)

def plan(store, **changes):
    return store.plan_update(changes or {'model': 'new-model'}, 'account:1')['plan_id']

def save(store):
    return store.confirm_update(plan(store), 'account:1', True)['save_id']

def test_mask_plan_snapshot_and_repr(store):
    before = sorted(store.root.iterdir())
    p = store.plan_update({'api_key': 'new-private-key'}, 'account:1')
    public = json.dumps([p, store.snapshot(), repr(store), repr(store._pending)])
    assert 'old-private-key' not in public and 'new-private-key' not in public
    assert p['settings']['api_key'] == '********'
    assert sorted(store.root.iterdir()) == before
    assert (store.root / '.env').read_bytes() == OLD

@pytest.mark.parametrize('changes', [None, {}, [], {'extra': 1}, {'provider': 'bad'}, {'provider': True}, {'context_window': True}, {'context_window': '32768'}, {'context_window': 1}, {'context_window': 2097153}, {'model': 'x\nINJECT=1'}, {'model': 'x\x00'}, {'model': '\ud800'}, {'model': 'x'*201}, {'api_key': 'x\nKEY=y'}, {'api_key': '${OTHER}'}, {'api_key': 'a=b'}, {'api_key': 'x'*4097}, {'embedding_channel': 'remote'}, {'embedding_model': '../a b'}, {'base_url': 'https://user:secret@example.test'}, {'base_url': 'https://example.test?token=secret'}, {'base_url': 'https://example.test#secret'}, {'base_url': 'file:///tmp/key'}, {'base_url': 'https://example.test:bad'}, {'base_url': 'http://example.test'}])
def test_schema_rejected_without_writes(store, changes):
    with pytest.raises(m.SettingsError) as err:
        store.plan_update(changes, 'account:1')
    assert str(err.value) == 'settings_invalid'
    assert (store.root / '.env').read_bytes() == OLD
    assert len(list(store.root.iterdir())) == 1

@pytest.mark.parametrize('approval', [False, 1, 'true', None, [], {}])
def test_confirmation_strict_consumed(store, approval):
    token = plan(store)
    r = store.confirm_update(token, 'account:1', approval)
    assert r['status'] == 'rejected' and r['wrote_nothing']
    assert store.confirm_update(token, 'account:1', True)['code'] == 'settings_plan_invalid'
    assert (store.root / '.env').read_bytes() == OLD

@pytest.mark.parametrize('kind', ['session', 'invalidate', 'expired', 'source_hash', 'source_stamp', 'source_identity', 'root'])
def test_stale_plan_consumed(store, monkeypatch, kind):
    token = plan(store)
    session = 'account:1'
    if kind == 'session':
        session = 'account:2'
    elif kind == 'invalidate':
        store.invalidate()
    elif kind == 'expired':
        store._pending[token].created -= 301
    elif kind == 'source_hash':
        (store.root / '.env').write_bytes(OLD.replace(b'old-model', b'bad-model'))
    elif kind == 'source_stamp':
        p = store.root / '.env'
        st = p.stat()
        os.utime(p, ns=(st.st_atime_ns, st.st_mtime_ns + 2000000000))
    elif kind == 'source_identity':
        p = store.root / '.env'
        p.rename(store.root / 'old.env')
        p.write_bytes(OLD)
    else:
        store._root_identity = (-1, -1)
    r = store.confirm_update(token, session, True)
    assert r['status'] == 'rejected' and r['wrote_nothing']
    assert store.confirm_update(token, 'account:1', True)['code'] == 'settings_plan_invalid'

@pytest.mark.parametrize('provider', ['deepseek', 'dashscope', 'ollama'])
def test_provider_mapping_preserve_comments_unknown(store, provider):
    changes = {'provider': provider, 'model': 'configured-model', 'base_url': 'http://localhost:11434' if provider == 'ollama' else 'https://example.test/v1', 'context_window': 32768, 'embedding_channel': 'api', 'embedding_model': 'text-embedding-v4'}
    if provider != 'ollama':
        changes['api_key'] = 'new-private-key'
    p = store.plan_update(changes, 'account:1')
    r = store.confirm_update(p['plan_id'], 'account:1', True)
    assert r['status'] == 'saved' and r['applied']
    data = (store.root / '.env').read_bytes()
    assert b'# preserved\r\n' in data and b'UNKNOWN_SETTING=value\r\n' in data
    assert (provider.upper() + '_MODEL=configured-model').encode() in data
    if provider == 'deepseek':
        assert b'# keep comment' in data
    snapshot = store.snapshot()
    assert snapshot['restart_required'] and snapshot['provider'] == provider
    assert snapshot['embedding_model'] == 'text-embedding-v4'
    directory = store._saved[r['save_id']].directory
    assert (directory / 'backup.env').read_bytes() == OLD
    for name in ['audit.json', 'audit-final.json']:
        audit = (directory / name).read_text('utf-8')
        assert 'private-key' not in audit and p['plan_id'] not in audit
        assert 'configured-model' not in audit

def test_secret_clear_and_keep(store):
    token = plan(store, api_key='')
    assert store.confirm_update(token, 'account:1', True)['applied']
    assert not store.snapshot()['api_key_configured']
    token = plan(store, model='other-model')
    assert store.confirm_update(token, 'account:1', True)['applied']
    assert not store.snapshot()['api_key_configured']

def test_ollama_key_rejected(store):
    with pytest.raises(m.SettingsError):
        plan(store, provider='ollama', api_key='secret')

def test_rollback_fresh_confirmation(store):
    saved = save(store)
    p = store.plan_rollback(saved, 'account:1')
    assert (store.root / '.env').read_bytes() != OLD
    denied = store.confirm_rollback(p['rollback_id'], 'account:1', False)
    assert denied['wrote_nothing']
    assert store.confirm_rollback(p['rollback_id'], 'account:1', True)['code'] == 'settings_plan_invalid'
    p = store.plan_rollback(saved, 'account:1')
    r = store.confirm_rollback(p['rollback_id'], 'account:1', True)
    assert r['restored'] and r['status'] == 'rolled_back'
    assert (store.root / '.env').read_bytes() == OLD

@pytest.mark.parametrize('target', ['source', 'backup', 'audit', 'directory'])
def test_rollback_refuses_changes(store, target):
    saved = save(store)
    p = store.plan_rollback(saved, 'account:1')
    record = store._saved[saved]
    if target == 'directory':
        record.identity = (-1, -1)
    else:
        path = store.root / '.env' if target == 'source' else record.directory / ('backup.env' if target == 'backup' else 'audit-final.json')
        path.write_bytes(b'foreign-edit')
    result = store.confirm_rollback(p['rollback_id'], 'account:1', True)
    assert result['status'] == 'rejected' and result['wrote_nothing']
    assert store.confirm_rollback(p['rollback_id'], 'account:1', True)['code'] == 'settings_plan_invalid'

def test_capacity_and_id_collision(store, monkeypatch):
    monkeypatch.setattr(m.secrets, 'token_hex', lambda n: 'a'*32)
    token = plan(store)
    with pytest.raises(m.SettingsError, match='settings_id_collision'):
        plan(store)
    store.confirm_update(token, 'account:1', False)
    with pytest.raises(m.SettingsError, match='settings_id_collision'):
        plan(store)
    monkeypatch.undo()
    for _ in range(m.MAX_PENDING):
        plan(store)
    with pytest.raises(m.SettingsError, match='settings_capacity'):
        plan(store)

def test_lifetime_cap(store):
    store._issued = {str(i) for i in range(m.MAX_IDS)}
    with pytest.raises(m.SettingsError, match='settings_capacity'):
        plan(store)

@pytest.mark.parametrize('data', [b'x'*65537, b'DEEPSEEK_MODEL=one\nDEEPSEEK_MODEL=two\n', b'DEEPSEEK_MODEL="multi\nline"\n', b'\xff', b'\x00', b'MODEL_PROVIDER=malicious\n', b'OFFICE_CONTEXT_WINDOW=raw-private-key\n'], ids=['oversize', 'duplicate', 'multiline', 'encoding', 'nul', 'provider', 'context'])
def test_bad_env_redacted(store, data):
    (store.root / '.env').write_bytes(data)
    with pytest.raises(m.SettingsError) as err:
        store.snapshot()
    assert str(err.value) in {'settings_file_invalid', 'settings_io_failed'}
    assert 'raw-private-key' not in str(err.value)

def test_no_json_migration(tmp_path):
    (tmp_path / 'settings.json').write_text('{"api_key":"legacy-secret"}')
    store = m.SettingsStore(tmp_path)
    assert not store.snapshot()['api_key_configured']
    assert store.snapshot()['migration'] == 'manual_reentry_only'
    assert len(list(tmp_path.iterdir())) == 1

def test_collision_not_overwritten_or_cleaned(store):
    token = plan(store)
    directory = store.root / ('.office-settings-' + token)
    directory.mkdir()
    (directory / 'foreign.txt').write_text('keep')
    result = store.confirm_update(token, 'account:1', True)
    assert result['status'] == 'uncertain' and not result['wrote_nothing']
    assert (directory / 'foreign.txt').read_text() == 'keep'
    assert (store.root / '.env').read_bytes() == OLD

def test_failure_after_replace_truthful_fused(store, monkeypatch):
    original = store._create
    def fail(path, data):
        if path.name == 'audit-final.json':
            raise OSError('new-private-key old-private-key')
        return original(path, data)
    monkeypatch.setattr(store, '_create', fail)
    result = store.confirm_update(plan(store), 'account:1', True)
    assert result['status'] == 'uncertain' and result['applied'] is None
    assert not result['wrote_nothing'] and result['restart_required']
    assert 'private-key' not in json.dumps(result)
    assert (store.root / '.env').read_bytes() != OLD
    with pytest.raises(m.SettingsError, match='settings_uncertain'):
        plan(store)

def test_external_edit_before_replace_retained(store, monkeypatch):
    original = store._create
    def changed(path, data):
        result = original(path, data)
        if path.name == 'replacement.env':
            (store.root / '.env').write_bytes(b'EXTERNAL=keep\n')
        return result
    monkeypatch.setattr(store, '_create', changed)
    result = store.confirm_update(plan(store), 'account:1', True)
    assert result['status'] == 'uncertain'
    assert (store.root / '.env').read_bytes() == b'EXTERNAL=keep\n'

def test_link_guard_without_link_privilege(store, monkeypatch):
    original = m._unsafe
    monkeypatch.setattr(m, '_unsafe', lambda st: st.st_size == len(OLD) or original(st))
    with pytest.raises(m.SettingsError, match='settings_file_invalid'):
        store.snapshot()

def test_invalidation_preserves_files_drops_rollback(store):
    saved = save(store)
    before = (store.root / '.env').read_bytes()
    store.invalidate()
    assert (store.root / '.env').read_bytes() == before
    with pytest.raises(m.SettingsError):
        store.plan_rollback(saved, 'account:1')


def test_old_and_new_secret_copies_are_masked(store):
    p = store.plan_update({'api_key': 'new-private-key', 'model': 'old-private-key'}, 'account:1')
    assert 'old-private-key' not in json.dumps(p)
    assert store.confirm_update(p['plan_id'], 'account:1', True)['applied']
    assert 'old-private-key' not in json.dumps(store.snapshot())
    p = store.plan_update({'model': 'new-private-key'}, 'account:1')
    assert 'new-private-key' not in json.dumps(p)


def test_unknown_credential_copy_masked(store):
    (store.root / '.env').write_bytes(OLD + b'OTHER_TOKEN=hidden-value\r\nDEEPSEEK_BASE_URL=https://example.test/hidden-value\r\n')
    assert 'hidden-value' not in json.dumps(store.snapshot())


def test_rollback_audit_failure_after_restoration_is_truthful(store, monkeypatch):
    saved = save(store)
    p = store.plan_rollback(saved, 'account:1')
    original = store._create
    def fail(path, data):
        if path.name == 'audit-final.json':
            raise OSError('raw-private-key')
        return original(path, data)
    monkeypatch.setattr(store, '_create', fail)
    result = store.confirm_rollback(p['rollback_id'], 'account:1', True)
    assert result['status'] == 'uncertain' and result['restored'] is None
    assert not result['wrote_nothing']
    assert 'raw-private-key' not in json.dumps(result)
    assert (store.root / '.env').read_bytes() == OLD


def test_owned_directory_replaced_before_backup_is_rejected(store, monkeypatch):
    original = store._create
    def changed(path, data):
        store._owned_dirs[path.parent] = (-1, -1)
        return original(path, data)
    monkeypatch.setattr(store, '_create', changed)
    result = store.confirm_update(plan(store), 'account:1', True)
    assert result['status'] == 'uncertain'
    assert (store.root / '.env').read_bytes() == OLD
    assert not list(store.root.glob('.office-settings-*/backup.env'))


def test_quoted_managed_value_preserves_comment(store):
    (store.root / '.env').write_bytes(OLD.replace(b'old-model # keep comment', b'"old-model" # keep comment'))
    assert store.snapshot()['model'] == 'old-model'
    assert store.confirm_update(plan(store), 'account:1', True)['applied']
    assert b'DEEPSEEK_MODEL=new-model # keep comment\r\n' in (store.root / '.env').read_bytes()


def test_absent_env_rollback_preview_explains_empty_restoration(tmp_path):
    store = m.SettingsStore(tmp_path)
    saved = save(store)
    p = store.plan_rollback(saved, 'account:1')
    assert p['restores_absent_source_as_empty_env'] is True
    assert store.confirm_rollback(p['rollback_id'], 'account:1', True)['restored']
    assert (tmp_path / '.env').read_bytes() == b''


@pytest.mark.parametrize('approval', [1, 'true', None])
def test_rollback_nonliteral_consumes(store, approval):
    saved = save(store)
    p = store.plan_rollback(saved, 'account:1')
    before = (store.root / '.env').read_bytes()
    r = store.confirm_rollback(p['rollback_id'], 'account:1', approval)
    assert r['status'] == 'rejected' and r['wrote_nothing']
    assert store.confirm_rollback(p['rollback_id'], 'account:1', True)['code'] == 'settings_plan_invalid'
    assert (store.root / '.env').read_bytes() == before


def test_hardlinked_source_rejected(store):
    os.link(store.root / '.env', store.root / 'linked.env')
    with pytest.raises(m.SettingsError, match='settings_file_invalid'):
        store.snapshot()
