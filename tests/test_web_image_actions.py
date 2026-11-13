"""P3 host image actions in the web bridge: fake service only, no provider/network/Office/GUI."""
import hashlib
import io
import json
import threading
from types import SimpleNamespace

import pytest
from PIL import Image

from office_agent.imagegen import ImageGenError
from office_agent.webui import bridge

PLAN = {"plan_id": "plan-1", "kind": "image_generation",
        "effective_prompt": "Requested visual style: natural.\nsynthetic private prompt",
        "size": "1024x1024", "style": "natural", "n": 1, "approval_ttl_seconds": 300.0,
        "configuration": {"provider": "openai_compatible", "model": "synthetic-model",
                          "base_url": "https://image.example.test/v1"}}


class MockAuth:
    def __init__(self, active=False):
        self.active = active

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
        return {"ok": False} if not confirmed else self.logout()


def png():
    stream = io.BytesIO()
    Image.new("RGB", (12, 8), "blue").save(stream, format="PNG")
    return stream.getvalue()


class FakeService:
    """Records every call; no plan/approval/audit logic and no network."""

    def __init__(self, *, error=None, block=None):
        self.calls = []
        self.invalidated = 0
        self.entered = threading.Event()
        self.raw = png()
        self.metadata = {"format": "PNG", "width": 12, "height": 8, "bytes": len(self.raw),
                         "sha256": hashlib.sha256(self.raw).hexdigest(), "metadata_stripped": True}
        self._error = error
        self._block = block

    def _log(self, name, args):
        self.calls.append((name, args))
        self.entered.set()
        if self._error is not None:
            raise self._error
        if self._block is not None:
            self._block.wait(timeout=5)

    def confirm_generation(self, plan_id, session_id, confirmed):
        self._log("confirm_generation", (plan_id, session_id, confirmed))
        return {"plan_id": plan_id, "status": "cancelled" if not confirmed else "generated",
                "transaction_id": plan_id, "image": dict(self.metadata)}

    def preview(self, plan_id, session_id):
        self.calls.append(("preview", (plan_id, session_id)))
        return SimpleNamespace(raw=self.raw, metadata=self.metadata)

    def plan_save(self, plan_id, session_id):
        self._log("plan_save", (plan_id, session_id))
        return {"save_id": "save-1", "path": "/authorized/office-image-save-1.png",
                "bytes": len(self.raw), "sha256": self.metadata["sha256"]}

    def confirm_save(self, save_id, session_id, confirmed):
        self._log("confirm_save", (save_id, session_id, confirmed))
        return {"save_id": save_id, "status": "saved", "written": True,
                "path": "/authorized/office-image-save-1.png"}

    def plan_rollback(self, save_id, session_id):
        self._log("plan_rollback", (save_id, session_id))
        return {"rollback_id": "rollback-1", "path": "/authorized/office-image-save-1.png",
                "sha256": self.metadata["sha256"]}

    def confirm_rollback(self, rollback_id, session_id, confirmed):
        self._log("confirm_rollback", (rollback_id, session_id, confirmed))
        return {"rollback_id": rollback_id, "status": "rolled_back", "removed": True,
                "path": "/authorized/office-image-save-1.png"}

    def invalidate(self):
        self.invalidated += 1

    def names(self):
        return [name for name, _ in self.calls]


def make_app(tmp_path, fake=None, session="session-1"):
    app = bridge.WebviewApp()
    app.path = tmp_path / "selected.docx"
    app.path.write_bytes(b"synthetic anchor - not an Office document")
    app.editor = SimpleNamespace(space=SimpleNamespace(root=tmp_path), audit=tmp_path / ".audit")
    if fake is not None:
        app._image_session = (session, fake)
    return app


def types(app):
    return [event["type"] for event in app.poll()]


def test_image_plan_event_is_display_only_and_never_persisted(tmp_path):
    app = make_app(tmp_path)
    directory = bridge._session_root(app.editor) / "images"
    directory.mkdir(parents=True)
    app._history_dir = directory
    app.emit("image_plan", PLAN)
    live = next(event for event in app.poll() if event["type"] == "image_plan")
    assert live["plan"]["plan_id"] == "plan-1"
    for expected in ("尺寸 1024x1024", "实际发送提示词", "费用未知且不可撤回", "独立确认", "另一次确认"):
        assert expected in live["text"]
    # The live card carries the id the UI needs; persisted history must not.
    assert bridge._display_event({"type": "image_plan", **live}) is None
    saved = bridge._read_history(directory)
    assert saved is not None and saved["events"] == []
    assert "plan-1" not in json.dumps(saved)


def test_strict_boolean_and_missing_session(tmp_path):
    fake = FakeService()
    app = make_app(tmp_path, fake)
    for consent in (1, 0, "true", "false", None, [], {}):
        assert app.confirm_image_generation("plan-1", consent)["reason"] == "confirmation_required"
    assert fake.calls == []
    assert app.confirm_image_save("save-1", "yes")["reason"] == "confirmation_required"
    assert app.confirm_image_rollback("rollback-1", 1)["reason"] == "confirmation_required"
    assert fake.calls == []
    empty = make_app(tmp_path)
    for call in (lambda: empty.confirm_image_generation("plan-1", True),
                 lambda: empty.plan_image_save("plan-1"),
                 lambda: empty.confirm_image_save("save-1", True),
                 lambda: empty.plan_image_rollback("save-1"),
                 lambda: empty.confirm_image_rollback("rollback-1", True)):
        assert call()["reason"] == "no_session"


def test_generation_returns_bounded_preview_never_raw(tmp_path):
    fake = FakeService()
    app = make_app(tmp_path, fake)
    result = app.confirm_image_generation("plan-1", True)
    assert result["ok"] and result["status"] == "generated" and result["preview_frames"] == 1
    assert fake.calls[0] == ("confirm_generation", ("plan-1", "session-1", True))
    assert not any(key in result for key in ("raw", "bytes", "data_url"))
    events = app.poll()
    card = next(event for event in events if event["type"] == "image_result")
    assert card["plan_id"] == "plan-1"
    assert len(card["images"]) == 1
    url = card["images"][0]["data_url"]
    assert url.startswith("data:image/png;base64,") and "未保存" in card["images"][0]["caption"]
    import base64
    assert base64.b64decode(url.split(",", 1)[1]) == fake.raw
    assert card["images"] == bridge._display_images([{"data_url": url, "alt": card["images"][0]["alt"],
                                                     "caption": card["images"][0]["caption"]}])
    assert not app.busy and app._busy_token is None
    # The unsaved preview is not history material either.
    assert bridge._display_event(card) is None


def test_image_badge_tracks_cancel_and_generated_preview(tmp_path):
    for confirmed, expected_label, expected_text in (
        (False, "生图已取消", "没有发出生图请求"),
        (True, "图片待保存", "图片已生成但尚未保存"),
    ):
        app = make_app(tmp_path, FakeService())
        app.emit("image_plan", PLAN)
        app.poll()
        result = app.confirm_image_generation("plan-1", confirmed)
        assert result["ok"]
        events = app.poll()
        assert next(event for event in events if event["type"] == "badge")["label"] == expected_label
        assert expected_text in next(event for event in events if event["type"] == "status")["text"]
        assert not app._image_plan_in_run


def test_three_plans_need_three_independent_decisions_and_status_tracks_remaining(tmp_path):
    fake = FakeService()
    app = make_app(tmp_path, fake)
    for index in range(1, 4):
        app.emit("image_plan", {**PLAN, "plan_id": f"plan-{index}",
                                "batch_index": index, "batch_total": 3})
    plans = [ev for ev in app.poll() if ev["type"] == "image_plan"]
    assert [ev["plan"]["plan_id"] for ev in plans] == ["plan-1", "plan-2", "plan-3"]
    assert app._image_pending_ids == {"plan-1", "plan-2", "plan-3"}
    assert fake.calls == []
    first = app.confirm_image_generation("plan-1", False)
    assert first["status"] == "cancelled"
    assert app._image_pending_ids == {"plan-2", "plan-3"}
    assert app._image_plan_in_run
    assert next(ev for ev in app.poll() if ev["type"] == "badge")["label"] == "等待生图确认"
    assert fake.names() == ["confirm_generation"]
    second = app.confirm_image_generation("plan-2", True)
    assert second["status"] == "generated"
    assert app._image_pending_ids == {"plan-3"}
    assert fake.names() == ["confirm_generation", "confirm_generation", "preview"]
    assert app._image_plan_in_run
    # Preview is NOT permission to save any image or to generate the last one.
    assert "plan_save" not in fake.names()
    assert app.confirm_image_generation("plan-3", False)["status"] == "cancelled"
    assert not app._image_pending_ids and not app._image_plan_in_run
    assert "plan_save" not in fake.names()


def test_image_badge_never_claims_unbilled_failure_after_attempt(tmp_path):
    app = make_app(tmp_path, FakeService(error=ImageGenError(
        "image_generation_failed", stage="response", request_state="attempt_committed")))
    app.emit("image_plan", PLAN)
    app.poll()
    response = app.confirm_image_generation("plan-1", True)
    assert response["provider_request_state"] == "attempt_committed"
    events = app.poll()
    assert next(event for event in events if event["type"] == "badge")["label"] == "生图结果待核对"
    text = next(event for event in events if event["type"] == "status")["text"]
    assert "可能计费" in text and "不要直接重复确认" in text


def test_generation_failure_only_returns_whitelisted_diagnostics(tmp_path):
    app = make_app(tmp_path, FakeService())
    def fail(service, session):
        raise ImageGenError("image_generation_failed", stage="response",
                            detail_code="image_response_invalid", request_state="attempt_committed")
    result = app._image_call(fail)
    assert result == {"ok": False, "reason": "image_error", "code": "image_generation_failed",
                      "failure_stage": "response", "failure_detail": "image_response_invalid",
                      "provider_request_state": "attempt_committed"}
    assert not app.busy
    def unsafe(service, session):
        raise ImageGenError("image_generation_failed", stage="private-details",
                            detail_code="private-key", request_state="other")
    assert app._image_call(unsafe) == {"ok": False, "reason": "image_error", "code": "image_generation_failed"}


def test_preview_failure_after_generation_preserves_possible_fee(tmp_path):
    class FailedPreview(FakeService):
        def preview(self, plan_id, session_id):
            raise ImageGenError("image_preview_unavailable")

    app = make_app(tmp_path, FailedPreview())
    result = app.confirm_image_generation("plan-1", True)
    assert result == {"ok": False, "reason": "image_error", "code": "image_preview_unavailable",
                      "provider_request_state": "attempt_committed"}
    assert "image_result" not in types(app)


def test_generation_cancel_consumes_without_preview(tmp_path):
    fake = FakeService()
    app = make_app(tmp_path, fake)
    result = app.confirm_image_generation("plan-1", False)
    assert result["ok"] and result["status"] == "cancelled"
    assert "provider_request_state" not in result
    assert fake.calls == [("confirm_generation", ("plan-1", "session-1", False))]
    assert "image_result" not in types(app)


def test_late_unscoped_idle_cannot_unlock_an_image_operation(tmp_path):
    release = threading.Event()
    fake = FakeService(block=release)
    app = make_app(tmp_path, fake)
    outcome = []
    worker = threading.Thread(target=lambda: outcome.append(app.confirm_image_generation("plan-1", True)))
    worker.start()
    try:
        assert fake.entered.wait(timeout=5)
        assert app.busy is True
        # A finished model task emits an unscoped idle while the image call owns the token.
        app.emit("idle", None)
        assert app.busy is True and app._busy_token is not None
    finally:
        release.set()
        worker.join(timeout=6)
    assert not worker.is_alive() and outcome[0]["ok"] is True
    assert app.busy is False and app._busy_token is None
    # With no token held, a normal idle still unlocks the window.
    app.emit("idle", None)
    assert app.busy is False


def test_session_fence_drops_a_late_generation_result(tmp_path):
    release = threading.Event()
    fake = FakeService(block=release)
    app = make_app(tmp_path, fake)
    outcome = []
    worker = threading.Thread(target=lambda: outcome.append(app.confirm_image_generation("plan-1", True)))
    worker.start()
    try:
        assert fake.entered.wait(timeout=5)
        with app._lock:  # simulates logout/selection clearing the bound session mid-flight
            app._image_session = None
    finally:
        release.set()
        worker.join(timeout=6)
    assert not worker.is_alive() and outcome[0]["reason"] == "stale"
    assert outcome[0]["provider_request_state"] == "attempt_committed"
    assert "image_result" not in types(app)
    assert app.busy is False and app._busy_token is None


def test_save_and_rollback_actions_report_their_own_cards(tmp_path):
    fake = FakeService()
    app = make_app(tmp_path, fake)
    save_plan = app.plan_image_save("plan-1")
    assert save_plan["ok"] and save_plan["result"]["save_id"] == "save-1"
    assert "不覆盖源文件" in next(e["text"] for e in app.poll() if e["type"] == "image_save_plan")
    saved = app.confirm_image_save("save-1", True)
    assert saved["ok"] and saved["result"]["status"] == "saved"
    saved_card = next(e for e in app.poll() if e["type"] == "image_saved")
    assert "已保存新文件" in saved_card["text"] and saved_card["save_id"] == "save-1"
    rollback = app.plan_image_rollback("save-1")
    assert rollback["ok"] and rollback["result"]["rollback_id"] == "rollback-1"
    plan_card = next(e for e in app.poll() if e["type"] == "image_rollback_plan")
    assert "不恢复历史" in plan_card["text"] and "不退费用" in plan_card["text"]
    removed = app.confirm_image_rollback("rollback-1", True)
    assert removed["ok"] and removed["result"]["removed"] is True
    assert "云端费用未撤销" in next(e["text"] for e in app.poll() if e["type"] == "image_rolled_back")
    assert fake.names() == ["plan_save", "confirm_save", "plan_rollback", "confirm_rollback"]
    cancels = app.confirm_image_save("save-1", False)
    assert cancels["ok"] is True and cancels["status"] == "cancelled"
    assert "image_saved" not in types(app)
    assert app.busy is False


def test_service_errors_are_mapped_to_stable_codes(tmp_path):
    for error, code in [(ImageGenError("image_approval_consumed"), "image_approval_consumed"),
                        (RuntimeError("synthetic-private-detail"), "image_failed")]:
        app = make_app(tmp_path, FakeService(error=error))
        result = app.confirm_image_generation("plan-1", True)
        assert result["ok"] is False and result["code"] == code
        if code == "image_failed":
            assert result["provider_request_state"] == "attempt_unknown"
        assert "synthetic-private-detail" not in json.dumps(result)
        assert app.busy is False and app._busy_token is None


def test_logout_fences_and_invalidates_the_service(tmp_path):
    fake = FakeService()
    app = make_app(tmp_path, fake)
    app._clear_private_state()
    assert fake.invalidated == 1 and app._image_session is None
    assert app.confirm_image_generation("plan-1", True)["reason"] == "no_session"
    other = FakeService()
    app._image_session = ("session-2", other)
    assert app._drop_image_session() is other
    assert app._image_session is None and app.poll() == []


def test_api_facade_requires_auth_and_never_coerces_truthy_consent(tmp_path):
    fake = FakeService()
    app = make_app(tmp_path, fake)
    api = bridge.Api(app, auth=MockAuth(active=False))
    calls = [lambda: api.confirm_image_generation("plan-1", True),
             lambda: api.plan_image_save("plan-1"),
             lambda: api.confirm_image_save("save-1", True),
             lambda: api.plan_image_rollback("save-1"),
             lambda: api.confirm_image_rollback("rollback-1", True)]
    for call in calls:
        assert call()["reason"] == "not_authenticated"
    assert fake.calls == []
    authenticated = bridge.Api(app, auth=MockAuth(active=True))
    assert authenticated.confirm_image_generation("plan-1", "true")["reason"] == "confirmation_required"
    assert fake.calls == []
    assert authenticated.confirm_image_generation("plan-1", True)["ok"] is True
    assert fake.calls[0] == ("confirm_generation", ("plan-1", "session-1", True))
    assert authenticated.plan_image_save("plan-1")["ok"] is True


@pytest.mark.parametrize("metadata", [{"format": "JPEG", "width": 12, "height": 8}, {}])
def test_non_png_or_incomplete_generated_metadata_yields_no_frame(tmp_path, metadata):
    fake = FakeService()
    fake.metadata = {**metadata, "sha256": "0" * 64, "bytes": 1}
    app = make_app(tmp_path, fake)
    result = app.confirm_image_generation("plan-1", True)
    assert result["ok"] and result["preview_frames"] == 0
    card = next(event for event in app.poll() if event["type"] == "image_result")
    assert "images" not in card or card["images"] == []
