"""The native P1 diagnostic must be self-contained and fail closed, without GUI in tests."""
import importlib
import json
from pathlib import Path
import runpy
import sys
from types import SimpleNamespace

import pytest

PROBE = Path(__file__).resolve().parents[1] / 'tools/startup_probe.py'


@pytest.fixture
def probe():
    return runpy.run_path(str(PROBE))


def test_hook_created_compiles_and_exact_owned_directory_is_removed(probe, tmp_path):
    directory = probe['prepare_hook'](tmp_path)
    assert directory.parent == tmp_path.resolve()
    assert directory.name.startswith(probe['HOOK_PREFIX'])
    assert (directory/'sitecustomize.py').read_text(encoding='utf-8') == probe['HOOK']
    compile(probe['HOOK'], 'sitecustomize.py', 'exec')
    assert probe['remove_hook'](directory, tmp_path) is True
    assert not directory.exists() and tmp_path.exists()


@pytest.mark.parametrize('nested', [False, True])
def test_cleanup_refuses_unowned_or_wrong_parent_directory(probe, tmp_path, nested):
    directory = tmp_path / ('nested' if nested else 'user-owned')
    directory.mkdir()
    if nested:
        directory = directory / (probe['HOOK_PREFIX']+'child')
        directory.mkdir()
    marker = directory/'keep.txt'
    marker.write_text('user owned', encoding='utf-8')
    with pytest.raises(RuntimeError, match='Refuse cleanup'):
        probe['remove_hook'](directory, tmp_path)
    assert marker.read_text(encoding='utf-8') == 'user owned'


def test_runtime_readiness_is_not_replaced_by_app_hook(probe, monkeypatch):
    monkeypatch.setattr(sys, 'orig_argv', ['python','-m','office_agent.webui.runtime'])
    monkeypatch.setattr(importlib, 'import_module', lambda *a: pytest.fail('runtime was intercepted'))
    exec(probe['HOOK'], {})


@pytest.mark.parametrize('args', [['--ui','tk'], ['--ui','web','--unexpected']])
def test_unexpected_app_arguments_exit_instead_of_unhooked_business_fallback(probe, monkeypatch, args):
    monkeypatch.setattr(sys, 'orig_argv', ['python','-m','office_agent.app',*args])
    monkeypatch.setattr(importlib, 'import_module', lambda *a: pytest.fail('unexpected app imported'))
    def terminate(code): raise SystemExit(code)
    import os
    monkeypatch.setattr(os, '_exit', terminate)
    with pytest.raises(SystemExit) as exc:
        exec(probe['HOOK'], {})
    assert exc.value.code == 97


def test_hook_install_exception_is_not_only_printed_and_ignored(probe, monkeypatch):
    monkeypatch.setattr(sys, 'orig_argv', ['python','-m','office_agent.app','--ui','web'])
    def failed_import(*args): raise RuntimeError('fixture import failure')
    monkeypatch.setattr(importlib, 'import_module', failed_import)
    def terminate(code): raise SystemExit(code)
    import os
    monkeypatch.setattr(os, '_exit', terminate)
    with pytest.raises(SystemExit) as exc:
        exec(probe['HOOK'], {})
    assert exc.value.code == 97


def install_fake(probe, monkeypatch):
    bridge = SimpleNamespace(WebviewApp=None, Api=None)
    monkeypatch.setattr(sys, 'orig_argv', ['python','-m','office_agent.app','--ui','web'])
    def fake_import(name):
        assert name == 'office_agent.webui.bridge'
        return bridge
    monkeypatch.setattr(importlib, 'import_module', fake_import)
    exec(probe['HOOK'], {})
    return bridge


def test_synthetic_controller_exposes_no_business_authentication_or_model_actions(probe, monkeypatch):
    bridge = install_fake(probe, monkeypatch)
    app = bridge.WebviewApp(ocr_provider=None)
    api = bridge.Api(app)
    assert api.auth_status() == {'ok':True, 'authenticated':False} and api.poll() == []
    for method in ('login','register','reset_accounts','select_file','submit_task','confirm_plan','llm'):
        assert not hasattr(api, method) and not hasattr(app, method)


def test_loaded_hook_reports_synthetic_dom_and_destroys_its_only_window(probe, monkeypatch, tmp_path):
    import threading
    class FakeThread:
        def __init__(self, **kwargs):
            assert kwargs['daemon'] is True
        def start(self): pass
    monkeypatch.setattr(threading, 'Thread', FakeThread)
    bridge = install_fake(probe, monkeypatch)
    target = tmp_path/'native.worker.json'
    monkeypatch.setenv('OFFICE_AGENT_STARTUP_PROBE_REPORT', str(target))
    class Event:
        def __init__(self): self.handlers = []
        def __iadd__(self, handler): self.handlers.append(handler); return self
    class Window:
        def __init__(self): self.events = SimpleNamespace(loaded=Event()); self.destroyed = 0
        def destroy(self): self.destroyed += 1
        def evaluate_js(self, script, callback):
            assert 'auth-screen' in script and 'document.readyState' in script
            callback({'ready':'complete','input':True,'css':True,'bridge':True,
                'auth_visible':True,'workspace_visible':False,'width':1000,'height':700})
    window = Window()
    bridge.WebviewApp().attach(window)
    assert len(window.events.loaded.handlers) == 1
    window.events.loaded.handlers[0]()
    result = json.loads(target.read_text(encoding='utf-8'))
    assert result['ok'] is True and result['hook_applied'] is True and result['synthetic_controller'] is True
    assert all(result[key] is False for key in ('task_submitted','document_selected','provider_called'))
    assert window.destroyed == 1
    window.events.loaded.handlers[0]()
    assert window.destroyed == 1  # Late/duplicate notifications cannot publish or destroy twice.


def test_prepare_write_failure_removes_exact_created_directory(probe, tmp_path, monkeypatch):
    original = Path.write_text
    def broken(path, *args, **kwargs):
        if path.name == 'sitecustomize.py':
            raise OSError('synthetic write failure')
        return original(path, *args, **kwargs)
    monkeypatch.setattr(Path, 'write_text', broken)
    with pytest.raises(OSError, match='synthetic write failure'):
        probe['prepare_hook'](tmp_path)
    assert list(tmp_path.iterdir()) == []
    assert probe['prepare_hook'].__globals__['_OWNED_HOOKS'] == {}


def test_foreign_prefix_matching_sibling_is_not_owned(probe, tmp_path):
    owned = probe['prepare_hook'](tmp_path)
    foreign = tmp_path / (probe['HOOK_PREFIX'] + 'foreign')
    foreign.mkdir()
    marker = foreign / 'keep.txt'
    marker.write_text('foreign', encoding='utf-8')
    with pytest.raises(RuntimeError, match='unregistered'):
        probe['remove_hook'](foreign, tmp_path)
    assert marker.read_text(encoding='utf-8') == 'foreign'
    assert probe['remove_hook'](owned, tmp_path)
    assert foreign.exists()


def test_replaced_registered_directory_identity_is_preserved(probe, tmp_path):
    owned = probe['prepare_hook'](tmp_path)
    # Keep the old directory live so its inode cannot be immediately recycled.
    moved = tmp_path / 'original-retained'
    assert owned.resolve() == owned and moved.resolve().parent == tmp_path.resolve()
    owned.rename(moved)
    owned.mkdir()
    marker = owned / 'keep.txt'
    marker.write_text('replacement', encoding='utf-8')
    with pytest.raises(RuntimeError, match='replaced'):
        probe['remove_hook'](owned, tmp_path)
    assert marker.read_text(encoding='utf-8') == 'replacement'
    assert (moved / 'sitecustomize.py').exists()


@pytest.mark.parametrize('method', ['is_symlink', 'is_junction'])
def test_link_or_junction_cleanup_refusal_without_os_privileges(probe, tmp_path, monkeypatch, method):
    owned = probe['prepare_hook'](tmp_path)
    original = getattr(Path, method, lambda self: False)
    with monkeypatch.context() as scoped:
        scoped.setattr(Path, method, lambda self: self == owned or original(self), raising=False)
        with pytest.raises(RuntimeError, match='Refuse cleanup'):
            probe['remove_hook'](owned, tmp_path)
        assert (owned / 'sitecustomize.py').exists()
    assert probe['remove_hook'](owned, tmp_path)


@pytest.mark.parametrize('method', ['is_symlink', 'is_junction'])
def test_link_or_junction_base_refused_before_creation(probe, tmp_path, monkeypatch, method):
    original = getattr(Path, method, lambda self: False)
    with monkeypatch.context() as scoped:
        scoped.setattr(Path, method, lambda self: self == tmp_path or original(self), raising=False)
        with pytest.raises(RuntimeError, match='base must not be a link'):
            probe['prepare_hook'](tmp_path)
    assert list(tmp_path.iterdir()) == []


@pytest.mark.parametrize('fails', [False, True])
def test_child_handle_close_records_result_and_allows_hook_cleanup_and_evidence(probe, tmp_path, fails):
    owned = probe['prepare_hook'](tmp_path)
    closes = []
    class Child:
        def close(self):
            closes.append('close')
            if fails:
                raise OSError('synthetic handle close failure')
    report = {'ok': True}
    probe['close_child_handle'](Child(), report)
    assert closes == ['close']
    assert report['child_handle_closed'] is (not fails)
    assert report['ok'] is (not fails)
    if fails:
        assert report['child_close_error'] == 'OSError: synthetic handle close failure'
    else:
        assert 'child_close_error' not in report
    report['hook_directory_removed'] = probe['remove_hook'](owned, tmp_path)
    evidence = tmp_path / 'evidence.json'
    evidence.write_text(json.dumps(report), encoding='utf-8')
    assert json.loads(evidence.read_text(encoding='utf-8')) == report
    assert report['hook_directory_removed'] is True and not owned.exists()
