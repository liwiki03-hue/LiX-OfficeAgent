"""Disposable XLSX selection -> read consent -> preview, without provider or COM."""
import hashlib
import json
from types import SimpleNamespace

from openpyxl import Workbook, load_workbook

from office_agent.agent import OfficeAgent
from office_agent.app import build_dispatcher, build_tools
from office_agent.webui import bridge


class OfflineModel:
    def __init__(self, name, args):
        self.name, self.args = name, args
        self.calls = 0

    def bind_tools(self, schemas):
        assert self.name in {item["function"]["name"] for item in schemas}
        return self

    def invoke(self, messages):
        self.calls += 1
        if self.calls == 1:
            return {"tool_calls": [{"id": "offline-only", "name": self.name, "args": self.args}]}
        return {"content": "Offline scripted response; inspect host evidence separately."}


def test_real_workbook_cancel_read_permission_and_write_preview(tmp_path, monkeypatch):
    source = tmp_path / "selected.xlsx"
    book = Workbook()
    sheet = book.active
    sheet.title = "Acceptance"
    sheet.append(["Label", "Amount"])
    sheet.append(["Synthetic", 7319])
    book.save(source)
    initial = hashlib.sha256(source.read_bytes()).hexdigest()
    chooser = []
    events = []
    app = bridge.WebviewApp()
    app.window = SimpleNamespace(create_file_dialog=lambda *a, **k: chooser[:])
    app._push = events.append
    # The bridge constructs a backend on selection, but never enters COM here.
    from office_agent import com
    monkeypatch.setattr(com, "WindowsComBackend", lambda **kwargs: object())

    assert app.select_file() == {"ok": False, "reason": "cancelled"}
    assert app.path is None and not events
    chooser.append(str(source))
    selected = app.select_file()
    assert selected["ok"] and selected["name"] == source.name
    assert app.path == source.resolve()
    assert events[-1]["type"] == "file"
    assert "7319" not in json.dumps(events[-1]["inspect"])
    assert hashlib.sha256(source.read_bytes()).hexdigest() == initial

    def run_read(label):
        model = OfflineModel("read_excel_region", {
            "path": str(source), "sheet": "Acceptance", "row": 2,
            "column": 1, "rows": 1, "columns": 2})
        allowed = [tool for tool in build_tools((source,)) if tool.name == model.name]
        agent = OfficeAgent(model, build_dispatcher(app.editor, source), allowed,
                            tmp_path / label, confirm=app._confirm_read, persist=False)
        state = agent.run(agent.new_state("Read the synthetic amount"))
        return agent, state

    def answer(event):
        events.append(event)
        if event["type"] == "read_confirm":
            app.answer_read(False, event["request_id"])

    app._push = answer
    refused_agent, refused = run_read("refused")
    assert refused_agent.results == {}
    assert len([event for event in events if event["type"] == "read_confirm"]) == 1
    assert app._read_authorized_path is None
    assert hashlib.sha256(source.read_bytes()).hexdigest() == initial

    def approve(event):
        events.append(event)
        if event["type"] == "read_confirm":
            app.answer_read(True, event["request_id"])

    app._push = approve
    agent, state = run_read("approved")
    assert state.status == "done" and not state.required.errors
    assert len(state.required.evidence) == 1
    locator = state.required.evidence[0].result_path
    result = json.loads(agent.results[locator.removeprefix("memory:")])
    assert any(cell["cell"] == "B2" and cell["value"] == 7319 for cell in result["cells"])
    assert app._read_authorized_path == source.resolve()
    prompt_count = len([event for event in events if event["type"] == "read_confirm"])
    run_read("approved-again")
    assert len([event for event in events if event["type"] == "read_confirm"]) == prompt_count

    # A write tool returns a plan only. Cancel/discard it, never call execute().
    plans = []
    dispatch = build_dispatcher(app.editor, source, on_plan=plans.append)
    preview = dispatch("plan_excel", {"path": str(source), "sheet": "Acceptance",
                                      "cell": "B2", "value": 8000})
    assert preview["plan_id"] and preview["preview"]["after"] == "8000"
    assert plans and app.execute()["reason"] == "no_plan"
    app.editor.plans.pop(preview["plan_id"])  # Discard the disposable preview; never execute it.
    assert hashlib.sha256(source.read_bytes()).hexdigest() == initial
    assert not list(tmp_path.glob("*.agent-*.xlsx"))

    chooser.clear()
    assert app.select_file()["reason"] == "cancelled"
    assert app.path == source.resolve() and app._read_authorized_path == source.resolve()
    chooser.append(str(source))
    assert app.select_file()["ok"]
    assert app._read_authorized_path is None
    verified = load_workbook(source, read_only=True)
    try:
        assert verified["Acceptance"]["B2"].value == 7319
    finally:
        verified.close()
