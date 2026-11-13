"""Focused checks for window-local read consent; no GUI, model or Office."""
from pathlib import Path
from types import SimpleNamespace
import threading

from office_agent.agent import Plan
from office_agent.webui import bridge


def setup_app(tmp_path):
    path = tmp_path / "selected.xlsx"
    path.write_bytes(b"test path only")
    app = bridge.WebviewApp()
    app.path = path.resolve()
    return app


def read_plan(app, tool="read_excel_region", access="read"):
    return Plan(tool=tool, access=access, arguments={"path": str(app.path)},
                file_versions={str(app.path): "test"})


def test_cancel_task_wakes_read_without_granting_permission(tmp_path):
    app = setup_app(tmp_path)
    assert app.cancel_task() == {"ok": False, "reason": "no_active_model_task"}
    event, decision = threading.Event(), [False]
    app.busy = True
    app._model_task_active = True
    app._pending_read = (read_plan(app), event, decision)
    assert app.cancel_task() == {"ok": True}
    assert app._stop_task.is_set() and event.is_set()
    assert decision == [False]


def test_one_consent_covers_multiple_current_file_read_tools(tmp_path):
    app = setup_app(tmp_path)
    prompts = []

    def approve(event):
        prompts.append(event)
        app.answer_read(True, event["request_id"])

    app._push = approve
    assert app._confirm_read(read_plan(app))
    assert app._confirm_read(read_plan(app, "read_excel_structure"))
    assert app._confirm_read(read_plan(app))
    assert len(prompts) == 1
    assert prompts[0]["scope"] == "selected_file_session"
    assert bridge.WebviewApp()._read_authorized_path is None


def test_cancel_preserves_but_successful_reselection_resets_consent(tmp_path, monkeypatch):
    app = setup_app(tmp_path)
    app._read_authorized_path = app.path
    selection = []
    app.window = SimpleNamespace(create_file_dialog=lambda *a, **k: selection)
    assert app.select_file()["reason"] == "cancelled"
    assert app._read_authorized_path == app.path

    from office_agent import com
    monkeypatch.setattr(com, "WindowsComBackend", lambda **kwargs: object())
    space = SimpleNamespace(root=tmp_path.resolve(), file=lambda value: Path(value))
    monkeypatch.setattr(bridge, "Editor", lambda *a, **k: SimpleNamespace(space=space))
    monkeypatch.setattr(bridge, "inspect_file", lambda path: {})
    selection.append(str(app.path))
    assert app.select_file()["ok"]
    assert app._read_authorized_path is None
    app._read_authorized_path = app.path
    other = tmp_path / "other.xlsx"
    other.write_bytes(b"test path only")
    selection[:] = [str(other)]
    assert app.select_file()["ok"]
    assert app.path == other.resolve()
    assert app._read_authorized_path is None


def test_batch_com_and_write_never_share_single_file_permission(tmp_path):
    app = setup_app(tmp_path)
    app._read_authorized_path = app.path
    prompts = []

    def approve(event):
        prompts.append(event)
        app.answer_read(True, event["request_id"])

    app._push = approve
    for tool in ("batch_documents", "com_read"):
        assert app._confirm_read(read_plan(app, tool))
        assert app._confirm_read(read_plan(app, tool))
    assert len(prompts) == 4
    assert all(event["scope"] == "once" for event in prompts)
    assert not app._confirm_read(read_plan(app, "plan_excel", access="write"))
    other = tmp_path / "other.xlsx"
    other.write_bytes(b"test path only")
    plan = Plan(tool="read_excel_region", access="read", arguments={"path": str(other)},
                file_versions={str(other): "test"})
    assert app._read_scope(plan) is None


def test_refusal_stale_and_expired_answers_do_not_grant_permission(tmp_path, monkeypatch):
    app = setup_app(tmp_path)

    def refuse(event):
        app.answer_read(True, "stale-request")
        assert app._read_authorized_path is None
        app.answer_read(False, event["request_id"])
        app.answer_read(True, event["request_id"])  # duplicate cannot reverse refusal

    app._push = refuse
    assert not app._confirm_read(read_plan(app))
    assert app._read_authorized_path is None

    def expire(event):
        app._pending_read_deadline = 0
        app.answer_read(True, event["request_id"])
        monkeypatch.setattr(app._pending_read[1], "wait", lambda timeout: False)

    app._push = expire
    assert not app._confirm_read(read_plan(app))
    assert app._pending_read is None
    assert app._read_authorized_path is None
