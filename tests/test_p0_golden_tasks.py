"""P0 golden task contracts: real host controller, deterministic offline workers.

No provider, Office COM, system chooser, billing, or existing business files are used.
These tests check the *order of visible events and authority*, not model quality.
"""
from types import SimpleNamespace

import pytest

from office_agent.agent import OfficeAgent
from office_agent.webui import bridge
from office_agent.writes import Workspace


@pytest.fixture
def harness(tmp_path, monkeypatch):
    path = tmp_path / "selected.xlsx"
    path.write_bytes(b"P0 isolated fixture; not an Office document")
    app = bridge.WebviewApp()
    app.path = path
    app.editor = SimpleNamespace(space=Workspace(tmp_path), audit=tmp_path / ".audit")
    starts = []

    class DormantWorker:
        def __init__(self, *, target, args, daemon):
            assert daemon and target == app._run_task
            self.args = args

        def start(self):
            starts.append(self.args)  # Do not contact a provider or start an OS process.

    monkeypatch.setattr(bridge.threading, "Thread", DormantWorker)
    return app, starts, path


# A user wording maps to either a read-only task, a planning task, a clarification
# card, or a rejection; *none* of these paths itself approves a write or a bill.
@pytest.mark.parametrize("case_id,question,outcome,family", [
    ("G01", "总结当前文档", "task", None),
    ("G02", "查看已有流程图", "task", None),
    ("G03", "帮我把表格改一下", "clarify", "write"),
    ("G04", "把B2改为7319", "task", None),
    ("G05", "请生成图表", "clarify", "diagram"),
    ("G06", "生成图片", "clarify", "image"),
    ("G07", "生成一张海边日落照片", "task", None),
    ("G08", "生成九张海边照片", "invalid_request", None),
    ("G09", "总结文档，然后把表格改一下", "clarify", "write"),
], ids=lambda value: value if isinstance(value, str) and value.startswith("G") else None)
def test_golden_entry_does_not_grant_authority(harness, case_id, question, outcome, family):
    app, starts, source = harness
    original = source.read_bytes()
    app.pending_id = "old-office-plan"  # A valid *new* task must discard this authority.
    response = app.ask(question)
    events = app.poll()
    if outcome == "invalid_request":
        assert response["reason"] == "invalid_request"
        assert app.pending_id == "old-office-plan" and not events and not starts
    else:
        assert response["ok"] and app.pending_id is None
        assert events[0] == {"type": "task", "text": question}
        cards = [event for event in events if event["type"] == "clarify"]
        if outcome == "clarify":
            assert response["clarification_required"] and len(cards) == 1
            assert cards[0]["questions"] and "不是" in cards[0]["preamble"]
            assert cards[0]["defaults"]["goal"] == "preview"
            assert cards[0]["classification"] in {"unclear", "multi-intent"}
            assert not starts and not app.busy
        else:
            assert not cards and len(starts) == 1 and app.busy
            assert starts[0][2] == question
        assert not app.execute()["ok"]  # No confirmation by merely asking/clarifying.
        assert not app._image_pending_ids and app._image_session is None
    assert source.read_bytes() == original, case_id
    assert not list(source.parent.glob("*.agent-*.xlsx"))


@pytest.mark.parametrize("case_id,question,goal,scope,preview", [
    ("G10", "帮我把表格改一下", "plan", "B2首尾空格", False),
    ("G11", "生成图片", "plan", "海边日落水彩插画", False),
    ("G12", "生成图表", "preview", "", True),
])
def test_golden_clarification_is_information_not_approval(harness, case_id, question, goal, scope, preview):
    app, starts, source = harness
    original = source.read_bytes()
    assert app.ask(question)["clarification_required"]
    card = next(event for event in app.poll() if event["type"] == "clarify")
    answers = {"goal": goal, "scope": scope} if scope else {"goal": goal}
    assert app.answer_clarification(card["request_id"], answers)["ok"]
    assert len(starts) == 1 and starts[0][3] is preview
    assert starts[0][5] == question  # Audit/UI preserve original user words.
    assert "不是授权" in starts[0][2]
    assert not app.answer_clarification(card["request_id"], answers)["ok"]
    assert not app.execute()["ok"] and not app._image_pending_ids
    assert source.read_bytes() == original, case_id


def test_golden_cancel_and_replace_question_revoke_old_clarification(harness):
    app, starts, source = harness
    original = source.read_bytes()
    assert app.ask("生成图片")["clarification_required"]
    old = next(event for event in app.poll() if event["type"] == "clarify")["request_id"]
    assert app.ask("总结当前文档")["ok"]
    assert not app.answer_clarification(old, {"goal": "plan", "scope": "海边日落"})["ok"]
    assert len(starts) == 1 and starts[0][2] == "总结当前文档"
    assert source.read_bytes() == original


@pytest.mark.parametrize("pending,failed,expected", [
    ("none", False, "模型已结束"),
    ("office", False, "等待确认"),
    ("image", False, "等待生图确认"),
    ("both", False, "等待确认"),
    ("image", True, "计划待核对"),
])
def test_golden_final_answer_does_not_claim_pending_effect(harness, pending, failed, expected):
    app, _, source = harness
    original = source.read_bytes()
    if pending in {"office", "both"}:
        app.emit("plan", {"id": "plan-not-approved", "kind": "synthetic", "confirmation_required": True})
    if pending in {"image", "both"}:
        app.emit("image_plan", {"plan_id": "image-not-approved", "approval_ttl_seconds": 120})
    app.emit("state", {"status": "done", "iterations": 1, "decision_limit": 32,
                       "token_estimate": 12, "context_capacity": 1000, "compressions": 0,
                       "answer": "请核对结果", "errors": ["offline error"] if failed else []})
    events = app.poll()
    assert [event["text"] for event in events if event["type"] == "answer"] == ["请核对结果"]
    assert next(event for event in events if event["type"] == "badge")["label"] == expected
    status = next(event for event in events if event["type"] == "status")["text"]
    if pending in {"office", "both"}:
        assert "尚未写入" in status
    if pending in {"image", "both"}:
        assert "尚未调用" in status and "服务" in status
    assert source.read_bytes() == original


@pytest.mark.parametrize("limit", [15, 32, 64])
def test_golden_decision_ceiling_has_exact_model_call_count(tmp_path, limit):
    class RepeatingInvalidModel:
        calls = 0

        def bind_tools(self, schemas):
            return self

        def invoke(self, messages):
            self.calls += 1
            return {"tool_calls": [{"name": "not_an_allowed_tool", "args": {}, "id": "invalid"}]}

    model = RepeatingInvalidModel()
    dispatches = []
    agent = OfficeAgent(model, lambda name, args: dispatches.append(name), [],
                        tmp_path / "checkpoint", decision_limit=limit, persist=False,
                        context_tokens=32768, reply_tokens=512, token_estimator=lambda _: 1)
    state = agent.run(agent.new_state("Synthetic no-op; bounded decisions"))
    assert state.status == "limit" and state.decisions == model.calls == limit
    assert not dispatches and any(f"{limit} model decisions" in error for error in state.required.errors)
