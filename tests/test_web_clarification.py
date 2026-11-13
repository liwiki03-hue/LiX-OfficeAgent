"""Preflight controller tests: no provider, native window, Office, or writes."""
import json
from types import SimpleNamespace
import pytest

from office_agent.webui import bridge
from office_agent.writes import Workspace


@pytest.fixture
def controller(tmp_path, monkeypatch):
    path = tmp_path / "fixture.txt"
    path.write_text("synthetic input", encoding="utf-8")
    app = bridge.WebviewApp()
    app.path = path
    app.editor = SimpleNamespace(space=Workspace(tmp_path))
    starts = []

    class Thread:
        def __init__(self, *, target, args, daemon):
            self.args = args
        def start(self):
            starts.append(self.args)

    monkeypatch.setattr(bridge.threading, "Thread", Thread)
    return app, starts


def card(app):
    return next(e for e in app.poll() if e["type"] == "clarify")


def test_preflight_runs_before_any_model_or_task_thread(controller):
    app, starts = controller
    app.pending_id = "old-plan"
    assert app.ask("帮我把表格改一下")["clarification_required"] is True
    event = card(app)
    assert len(event["questions"]) == 2
    assert not starts and not app.busy and app.pending_id is None
    assert app._read_authorized_path is None
    assert not app.execute()["ok"]
    assert app.answer_clarification(event["request_id"], {"goal": "plan", "scope": "B2首尾空格"})["ok"]
    assert len(starts) == 1 and app.busy
    assert "B2首尾空格" in starts[0][2]
    assert "不是授权" in starts[0][2]
    assert app._read_authorized_path is None
    assert app.pending_id is None
    app.busy = False
    assert not app.answer_clarification(event["request_id"], {})["ok"]
    assert len(starts) == 1


def test_unbounded_image_count_rejected_before_task_or_old_session_reset(controller):
    app, starts = controller
    app._image_pending_ids.add("old-plan")
    assert app.ask("生成九张海边的照片")["reason"] == "invalid_request"
    assert not starts and app._image_pending_ids == {"old-plan"}
    assert app.ask("生成一张1024x1024图片")["clarification_required"]
    assert not starts


def test_image_count_in_user_clarification_answer_is_host_bounded(controller):
    app, starts = controller
    assert app.ask("生成图片")["clarification_required"]
    request_id = card(app)["request_id"]
    assert app.answer_clarification(request_id, {"goal": "plan", "scope": "画十张海边的照片"})["reason"] == "invalid_answer"
    assert not starts  # correction is allowed; no batch, billing, or agent started
    assert app.answer_clarification(request_id, {"goal": "plan", "scope": "画三张海边的照片"})["ok"]
    assert len(starts) == 1 and starts[0][6]["image_count"] == 3
    assert starts[0][5] == "生成图片"  # UI/history retain the original question
    assert "三张" in starts[0][2]  # model receives the user's additional subject


def test_skip_records_exact_visible_defaults(controller):
    app, starts = controller
    app.ask("生成图表")
    event = card(app)
    defaults = dict(event["defaults"])
    # Public payload mutation cannot modify the stored defaults.
    event["defaults"]["goal"] = "approved"
    assert app.answer_clarification(event["request_id"], {}, True)["ok"]
    task = starts[0][2]
    assert '"skipped": true' in task
    assert '"goal": "preview"' in task
    assert '"defaults_used"' in task
    assert all(value in task for value in defaults.values())
    assert app.pending_id is None and app._read_authorized_path is None


def test_read_task_never_asks(controller):
    app, starts = controller
    assert app.ask("总结文档") == {"ok": True}
    assert len(starts) == 1
    assert not any(e["type"] == "clarify" for e in app.poll())


@pytest.mark.parametrize("change", ["new_task", "logout", "file_changed", "path_changed", "expired"])
def test_stale_answer_never_starts_task(controller, monkeypatch, change):
    app, starts = controller
    app.ask("帮我把表格改一下")
    old = card(app)["request_id"]
    if change == "new_task":
        app.ask("生成图表")
    elif change == "logout":
        app._clear_private_state()
    elif change == "file_changed":
        app.path.write_text("changed", encoding="utf-8")
    elif change == "path_changed":
        app.path = app.path.with_name("other.txt")
    else:
        monkeypatch.setattr(bridge.time, "monotonic", lambda: 10**20)
    assert not app.answer_clarification(old, {}, True)["ok"]
    assert not starts and not app.busy


def test_invalid_answer_can_be_corrected_without_starting(controller):
    app, starts = controller
    app.ask("生成图表")
    request_id = card(app)["request_id"]
    assert not app.answer_clarification(request_id, {"confirmed": True})["ok"]
    assert not starts
    assert app.answer_clarification(request_id, {}, True)["ok"]
    assert len(starts) == 1


def test_answer_api_requires_authentication(controller):
    app, starts = controller
    app.ask("生成图表")
    request_id = card(app)["request_id"]
    auth = SimpleNamespace(status=lambda: {"authenticated": False})
    api = bridge.Api(app, auth=auth)
    assert api.answer_clarification(request_id, {}, True)["reason"] == "not_authenticated"
    assert not starts


def test_defaults_persist_as_goal_information_not_approvals(controller, monkeypatch):
    app, starts = controller
    app.editor.audit = app.editor.space.root / "owned-audit"
    app.ask("生成图表")
    request_id = card(app)["request_id"]
    assert app.answer_clarification(request_id, {}, True)["ok"]
    def no_provider():
        raise RuntimeError("intentional provider-free test stop")
    monkeypatch.setattr(app, "_build_model", no_provider)
    app._run_task(*starts[0])
    directories = list((app.editor.audit / "sessions").iterdir())
    assert len(directories) == 1
    metadata = json.loads((directories[0] / "session.json").read_text(encoding="utf-8"))
    history = json.loads((directories[0] / "events.json").read_text(encoding="utf-8"))
    # History shows the user's wording; the machine-readable decision is audit-only.
    assert metadata["question"] == "生成图表"
    assert '"skipped": true' in metadata["goal"]
    assert '"goal": "preview"' in metadata["goal"]
    assert '"defaults_used"' in metadata["goal"]
    assert any('"defaults_used"' in e.get("text", "") for e in history["events"])
    assert not ({"approved", "confirmed", "plan_id"} & metadata.keys())
    assert app.pending_id is None and app._read_authorized_path is None


def test_skipped_suggestions_cannot_unlock_write_or_generation(controller, monkeypatch):
    app, starts = controller
    app.editor.audit = app.editor.space.root / "owned-audit"
    app.ask("生成图表")
    assert app.answer_clarification(card(app)["request_id"], {}, True)["ok"]
    assert starts[0][3] is True
    monkeypatch.setattr(app, "_build_model", lambda: object())
    monkeypatch.setattr(bridge._app, "classify_document", lambda _: SimpleNamespace(kind="word"))
    raw_calls, observations = [], []
    monkeypatch.setattr(bridge._app, "build_dispatcher", lambda *a, **kw: lambda name, args: raw_calls.append(name))
    def observe_agent(model, dispatch, tools, state_dir, **kwargs):
        observations.append(kwargs)
        assert kwargs["write_tools"] == [] and kwargs["write_gate"] is None
        assert all(t.access == "read" for t in tools)
        assert all(t.access == "read" for t in kwargs["capability_tools"])
        for name in ("image_gen", "plan_word", "request_write", bridge._app._WRITE_GATE.name, "unknown"):
            with pytest.raises(ValueError, match="仅允许只读"):
                dispatch(name, {})
        observations.append("all_guards_passed")
        raise RuntimeError("intentional provider-free observation stop")
    monkeypatch.setattr(bridge, "OfficeAgent", observe_agent)
    app._run_task(*starts[0])
    assert len(observations) == 2 and observations[-1] == "all_guards_passed"
    assert raw_calls == []
    assert app.pending_id is None and app._image_session is None
