"""Focused startup/freeze checks: no GUI, CLR, model service or Office process."""
import sys
from types import SimpleNamespace

import pytest

from office_agent import app
from office_agent.webui import runtime


@pytest.mark.parametrize("platform,gui,expected", [
    ("win32", None, "edgechromium"),
    ("linux", None, "qt"),
    ("win32", "qt", "qt"),
])
def test_backend_selected_before_import_and_explicit_runtime_preserved(tmp_path, monkeypatch,
                                                                       platform, gui, expected):
    monkeypatch.setattr(runtime, "PROJECT_ROOT", tmp_path)
    monkeypatch.setattr(sys, "platform", platform)
    monkeypatch.delenv("OFFICE_WEBVIEW_GUI", raising=False)
    monkeypatch.delenv("PYTHONNET_RUNTIME", raising=False)
    if gui:
        monkeypatch.setenv("OFFICE_WEBVIEW_GUI", gui)
    assert runtime.configure_runtime() == expected
    import os
    assert os.getenv("PYTHONNET_RUNTIME") == ("netfx" if expected == "edgechromium" else None)
    monkeypatch.setenv("PYTHONNET_RUNTIME", "coreclr")
    assert runtime.configure_runtime() == expected
    assert os.getenv("PYTHONNET_RUNTIME") == "coreclr"


@pytest.mark.parametrize("value,expected", [
    (None, None), ("", None), ("abc", None), ("0", None), ("99999", None),
    ("-1", None), ("+9222", None), (" 9222", None), ("1.0", None),
    ("１２３４", None), ("1", 1), ("9222", 9222), ("65535", 65535),
])
def test_webview_debug_port_is_explicit_and_strict(tmp_path, monkeypatch, value, expected):
    from office_agent.webui import bridge

    monkeypatch.setattr(runtime, "configure_runtime", lambda: "qt")
    monkeypatch.setattr(bridge, "WebviewApp", lambda **kwargs: SimpleNamespace(attach=lambda window: None))
    monkeypatch.setattr(bridge, "Api", lambda target: object())
    fake = SimpleNamespace(settings={"REMOTE_DEBUGGING_PORT": None},
                           create_window=lambda *a, **kw: object(),
                           start=lambda **kw: None)
    monkeypatch.setitem(sys.modules, "webview", fake)
    if value is None:
        monkeypatch.delenv("OFFICE_WEBVIEW_DEBUG_PORT", raising=False)
    else:
        monkeypatch.setenv("OFFICE_WEBVIEW_DEBUG_PORT", value)
    app._launch_web()
    assert fake.settings["REMOTE_DEBUGGING_PORT"] == expected
    if expected is not None:
        # Do not retain CDP access after an earlier opt-in in this process.
        monkeypatch.delenv("OFFICE_WEBVIEW_DEBUG_PORT", raising=False)
        app._launch_web()
        assert fake.settings["REMOTE_DEBUGGING_PORT"] is None


def test_window_creation_failure_reports_original_error_and_safe_fallback(tmp_path, monkeypatch):
    monkeypatch.setenv('OFFICE_FLOATING_ORB', '0')
    monkeypatch.setattr(runtime, "PROJECT_ROOT", tmp_path)
    monkeypatch.setattr(sys, "platform", "win32")
    monkeypatch.setenv("OFFICE_WEBVIEW_GUI", "edgechromium")
    monkeypatch.delenv("PYTHONNET_RUNTIME", raising=False)

    def fail_create(*args, **kwargs):
        import os
        assert os.getenv("PYTHONNET_RUNTIME") == "netfx"
        raise RuntimeError("synthetic loader failure")

    monkeypatch.setitem(sys.modules, "webview", SimpleNamespace(
        settings={"REMOTE_DEBUGGING_PORT": None}, create_window=fail_create))
    with pytest.raises(SystemExit) as exc:
        app._launch_web()
    assert "synthetic loader failure" in str(exc.value)
    assert "Start-OfficeAgent.ps1" in str(exc.value)
    assert "uv pip install PySide6" not in str(exc.value)


def test_windows_edge_default_uses_floating_host_and_allows_diagnostic_opt_out(monkeypatch):
    import os
    if os.name != 'nt':
        pytest.skip('Windows native host selection')
    from office_agent.webui import floating_shell, bridge
    calls = []
    monkeypatch.setattr(runtime, 'configure_runtime', lambda: 'edgechromium')
    monkeypatch.setattr(floating_shell, 'launch_floating', lambda *args: calls.append('floating'))
    monkeypatch.setattr(bridge, 'WebviewApp', lambda **kwargs: SimpleNamespace(attach=lambda window: None))
    monkeypatch.setattr(bridge, 'Api', lambda target: object())
    monkeypatch.setitem(sys.modules, 'webview', SimpleNamespace(settings={},
        create_window=lambda *args, **kwargs: calls.append('diagnostic'), start=lambda **kwargs: None))
    monkeypatch.delenv('OFFICE_FLOATING_ORB', raising=False)
    app._launch_web()
    monkeypatch.setenv('OFFICE_FLOATING_ORB', '0')
    app._launch_web()
    assert calls == ['floating', 'diagnostic']


@pytest.mark.parametrize("name,suffix", [("plan_legacy_word_replace", ".doc"),
                                         ("plan_legacy_cell_set", ".xls")])
def test_frozen_preview_never_calls_editor_or_classifier(tmp_path, monkeypatch, name, suffix):
    path = tmp_path / ("selected" + suffix)
    path.write_bytes(b"synthetic bytes; never passed to Office")

    def forbidden(*args, **kwargs):
        raise AssertionError("Frozen preview must not inspect/plan/activate COM")

    editor = SimpleNamespace(space=SimpleNamespace(file=lambda value: path), plan=forbidden)
    monkeypatch.setattr(app, "classify_document", forbidden)
    dispatch = app.build_dispatcher(editor, path)
    with pytest.raises(ValueError, match="已冻结") as exc:
        dispatch(name, {"path": str(path)})
    assert "未做任何写入" in app.legacy_refusal_presentation(exc.value)
    assert path.read_bytes() == b"synthetic bytes; never passed to Office"
    assert list(tmp_path.iterdir()) == [path]
