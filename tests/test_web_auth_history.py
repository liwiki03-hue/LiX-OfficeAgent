"""Deterministic auth/history boundaries: no GUI, model or Office."""
import json
import os
import threading
from pathlib import Path
from types import SimpleNamespace

from office_agent.webui import bridge


class MockAuth:
    def __init__(self):
        self.active = False

    def status(self):
        return {"ok": True, "authenticated": self.active, "user": "local" if self.active else None}

    def register(self, username, password):
        return self.status()

    def login(self, username, password):
        self.active = True
        return self.status()

    def logout(self):
        self.active = False
        return self.status()

    def reset_accounts(self, confirmed=False):
        if not confirmed:
            return {"ok": False}
        return self.logout()


def setup_app(tmp_path):
    app = bridge.WebviewApp()
    app.path = tmp_path / "selected.xlsx"
    app.path.write_bytes(b"fake selected file")
    app.editor = SimpleNamespace(space=SimpleNamespace(root=tmp_path), audit=tmp_path / ".audit")
    return app


def session(app, name="one", events=None, truncated=False, selected=None):
    directory = bridge._session_root(app.editor) / name
    directory.mkdir(parents=True)
    (directory / "session.json").write_text(json.dumps({"version": 1, "selected_path": str(selected or app.path), "question": "question"}), encoding="utf-8")
    if events is not None:
        (directory / "events.json").write_text(json.dumps({"version": 1, "events": events, "truncated": truncated}), encoding="utf-8")
    return directory


def test_unauthenticated_never_dispatches(monkeypatch):
    app = bridge.WebviewApp()
    api = bridge.Api(app, auth=MockAuth())
    def forbidden(*args):
        raise AssertionError("Business called without authentication")
    for name in ("poll", "snapshot", "list_sessions", "select_file", "ask", "execute", "undo", "probe", "answer_read", "load_session"):
        monkeypatch.setattr(app, name, forbidden)
        args = ("value",) if name in {"ask", "load_session"} else (True,) if name == "answer_read" else ()
        assert getattr(api, name)(*args)["reason"] == "not_authenticated"
    assert api.resume()["reason"] == "not_authenticated"
    assert not api.auth_status()["authenticated"]
    api.register("local", "password")
    assert not api.auth_status()["authenticated"]


def test_login_logout_clear_old_authorizations(tmp_path):
    app = setup_app(tmp_path)
    auth = MockAuth()
    api = bridge.Api(app, auth=auth)
    app._read_authorized_path = app.path
    app.pending_id = app.last_verified_id = "old"
    app._push({"type": "log", "text": "private"})
    assert api.login("local", "password")["ok"]
    assert app.path is None and app.editor is None
    assert app._read_authorized_path is None and app.pending_id is None
    event = threading.Event()
    decision = [True]
    app._pending_read = (None, event, decision)
    app.events.put(("text", "private"))
    assert api.logout()["ok"]
    assert event.is_set() and not decision[0]
    assert app.events.empty() and not app._buffer
    assert api.poll()["reason"] == "not_authenticated"


def test_busy_logout_reset_refuse_without_changing_account():
    app = bridge.WebviewApp()
    auth = MockAuth()
    api = bridge.Api(app, auth=auth)
    api.login("local", "password")
    app.busy = True
    assert api.logout()["reason"] == "busy"
    assert api.reset_accounts(True)["reason"] == "busy"
    assert auth.active
    app.busy = False
    assert not api.reset_accounts()["ok"]
    assert api.reset_accounts(True)["ok"] and not auth.active


def test_load_scope_legacy_truncated_and_readonly(tmp_path):
    app = setup_app(tmp_path)
    legacy = session(app)
    partial = session(app, "partial", [{"type": "task", "text": "saved"}], True)
    other = session(app, "other", [], selected=tmp_path / "other.xlsx")
    assert len(app.list_sessions()["sessions"]) == 2
    assert app.load_session(str(legacy))["history_available"] is False
    assert "没有保存" in app.load_session(str(legacy))["warning"]
    app._read_authorized_path = app.path
    app.pending_id = "current"
    result = app.load_session(str(partial))
    assert result["readonly"] and result["history_truncated"]
    assert result["events"] == [{"type": "task", "text": "saved"}]
    assert app.pending_id == "current" and app._read_authorized_path == app.path
    assert not app.load_session(str(other))["ok"]
    assert not app.load_session(str(tmp_path))["ok"]


def test_history_scans_newest_metadata_before_bounded_read(tmp_path, monkeypatch):
    app = setup_app(tmp_path)
    root = bridge._session_root(app.editor)
    older = []
    for index in range(500):
        directory = session(app, f"older-{index:03d}")
        stamp = 1_700_000_000_000_000_000 + index * 1_000_000
        os.utime(directory / "session.json", ns=(stamp, stamp))
        older.append(directory)
    latest = session(app, "latest")
    stamp = 1_700_000_002_000_000_000
    os.utime(latest / "session.json", ns=(stamp, stamp))
    original = Path.iterdir
    monkeypatch.setattr(Path, "iterdir", lambda path: iter([*older, latest]) if path == root else original(path))
    listing = app.list_sessions()
    assert listing["ok"] and len(listing["sessions"]) == 100
    assert listing["sessions"][0]["session_dir"] == str(latest.resolve())
    assert "500" in listing["warning"]
    replay = app.load_session(str(latest))
    assert replay["ok"] and replay["readonly"]


def test_image_plan_final_status_awaits_explicit_paid_approval():
    app = bridge.WebviewApp()
    app.emit("image_plan", {"plan_id": "offline-only", "approval_ttl_seconds": 120})
    app.emit("state", {"status": "done", "iterations": 1, "decision_limit": 32,
                       "tokens": 10, "capacity": 1000, "compressions": 0,
                       "answer": "请确认生图", "errors": []})
    events = app.poll()
    assert any(event["type"] == "image_plan" for event in events)
    assert next(event for event in events if event["type"] == "badge")["label"] == "等待生图确认"
    assert "尚未调用生图服务" in next(event for event in events if event["type"] == "status")["text"]
    app.emit("state", {"status": "done", "iterations": 1, "decision_limit": 32,
                       "tokens": 10, "capacity": 1000, "compressions": 0,
                       "answer": "", "errors": ["offline test"]})
    events = app.poll()
    assert next(event for event in events if event["type"] == "badge")["label"] == "计划待核对"
    assert "生图计划尚未调用服务" in next(event for event in events if event["type"] == "status")["text"]


def test_links_and_malicious_history_rejected(tmp_path, monkeypatch):
    app = setup_app(tmp_path)
    directory = session(app, events=[{"type": "plan", "plan": {"id": "executable"}}])
    assert not app.load_session(str(directory))["ok"]
    original = Path.is_symlink
    monkeypatch.setattr(Path, "is_symlink", lambda self: self == directory or original(self))
    assert not app.load_session(str(directory))["ok"]


def test_history_is_bounded_sanitized_and_poll_not_duplicated(tmp_path, monkeypatch):
    app = setup_app(tmp_path)
    directory = session(app)
    app._history_dir = directory
    app._push({"type": "task", "text": "first"})
    app._push({"type": "read_confirm", "request_id": "never-save"})
    app.emit("plan", {"id": "execute-secret", "kind": "test"})
    app.emit("result", {"tool": "read", "text": "authorized", "data": {"value": 1, "request_id": "secret"}})
    before = bridge._read_history(directory)
    assert len(before["events"]) == 3
    assert "execute-secret" not in json.dumps(before)
    assert before["events"][-1]["data"] == {"value": 1}
    app.poll()
    app.poll()
    assert bridge._read_history(directory) == before
    monkeypatch.setattr(bridge, "_HISTORY_COUNT", 3)
    app._push({"type": "log", "text": "too many"})
    saved = bridge._read_history(directory)
    assert saved["truncated"] and len(saved["events"]) == 3


def test_final_answer_is_separate_replayable_event(tmp_path):
    app = setup_app(tmp_path)
    directory = session(app)
    app._history_dir = directory
    app._handle_event("state", {"status": "done", "iterations": 17,
        "decision_limit": 32, "token_estimate": 45, "context_capacity": 1000,
        "compressions": 0, "answer": "这是最终答复", "errors": [],
        "resume_checkpoint": "private-checkpoint"})
    saved = bridge._read_history(directory)["events"]
    assert [event["text"] for event in saved if event["type"] == "answer"] == ["这是最终答复"]
    assert next(event for event in saved if event["type"] == "progress")["decision_limit"] == 32
    assert all("private-checkpoint" not in json.dumps(event, ensure_ascii=False) for event in saved)
    assert all("这是最终答复" not in event.get("text", "") for event in saved if event["type"] == "log")
    assert app.load_session(str(directory))["events"] == saved


def test_real_task_seeds_first_event_before_model_failure(tmp_path, monkeypatch):
    app = setup_app(tmp_path)
    monkeypatch.setattr(app, "_build_model", lambda: (_ for _ in ()).throw(RuntimeError("offline")))
    app._run_task(app.path, app.editor, "first question")
    listing = app.list_sessions()
    assert len(listing["sessions"]) == 1
    directory = Path(listing["sessions"][0]["session_dir"])
    assert directory.parent == bridge._session_root(app.editor)
    saved = bridge._read_history(directory)
    assert saved["events"][0] == {"type": "task", "text": "first question"}
    assert any("offline" in e.get("text", "") for e in saved["events"])
    assert not app.busy


def test_size_limit_and_known_credentials_redacted(tmp_path, monkeypatch):
    app = setup_app(tmp_path)
    directory = session(app)
    app._history_dir = directory
    monkeypatch.setenv("EXAMPLE_API_KEY", "private-key-value")
    app._push({"type": "result", "text": "private-key-value", "data": {"password": "secret", "value": 1}})
    saved = bridge._read_history(directory)
    assert "private-key-value" not in json.dumps(saved)
    assert saved["events"][0]["data"] == {"value": 1}
    app._push({"type": "log", "text": "结果：未完成\n可续跑检查点：private/agent-state.json\n重新提交"})
    assert "private/agent-state.json" not in json.dumps(bridge._read_history(directory))
    monkeypatch.setattr(bridge, "_HISTORY_LIMIT", 1024)
    app._push({"type": "log", "text": "x" * 2000})
    saved = bridge._read_history(directory)
    assert saved["truncated"] and len(saved["events"]) == 2


def test_existing_history_is_not_overwritten_warn_once(tmp_path):
    app = setup_app(tmp_path)
    directory = session(app, events=[{"type": "task", "text": "unowned"}])
    original = (directory / "events.json").read_bytes()
    app._history_dir = directory
    app._push({"type": "task", "text": "new"})
    app._push({"type": "task", "text": "again"})
    assert (directory / "events.json").read_bytes() == original
    assert len([e for e in app.poll() if e.get("tag") == "warn"]) == 1


def test_final_state_reports_uncertain_outcome_without_changing_state_contract():
    app = bridge.WebviewApp()
    payload = {"status": "done", "iterations": 3, "token_estimate": 120,
               "context_capacity": 2000, "compressions": 0, "answer": "请查看结果",
               "errors": ["Tool failed: image_configuration_missing"],
               "resume_checkpoint": None}
    app.emit("state", payload)
    events = app.poll()
    assert next(e for e in events if e["type"] == "status")["text"].startswith("处理结束，但过程中有步骤报错")
    assert next(e for e in events if e["type"] == "badge")["label"] == "结果待核对"
    assert next(e for e in events if e["type"] == "log" and "结果：" in e["text"])["tag"] != "ok"
    assert payload["status"] == "done" and len(payload["errors"]) == 1

    app.pending_id = "pending-plan"
    app.emit("state", {**payload, "errors": []})
    events = app.poll()
    assert next(e for e in events if e["type"] == "status")["text"].startswith("修改计划已生成，尚未写入")
    assert next(e for e in events if e["type"] == "badge")["label"] == "等待确认"
    assert app.pending_id == "pending-plan"

    app.emit("state", payload)
    events = app.poll()
    assert next(e for e in events if e["type"] == "badge")["label"] == "计划待核对"
    app.pending_id = None
    app.emit("state", {**payload, "errors": []})
    events = app.poll()
    assert next(e for e in events if e["type"] == "badge")["label"] == "模型已结束"
    assert next(e for e in events if e["type"] == "status")["level"] == "info"
