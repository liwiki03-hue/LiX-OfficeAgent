"""Offline end-to-end preview boundary from a mocked model to a verified output."""
from docx import Document

from office_agent.agent import OfficeAgent, ToolSpec
from office_agent.writes import Editor, WordReplace, digest


class ScriptedModel:
    def __init__(self, path):
        self.responses = iter([
            {"tool_calls": [{"name": "plan_word", "args": {
                "kind": "word_replace", "path": str(path), "old": "cost", "new": "price",
                "paragraph_index": 0, "expected_count": 1}, "id": "one"}]},
            {"content": "Plan prepared; user must separately approve."},
        ])

    def bind_tools(self, schemas):
        return self

    def invoke(self, messages):
        return next(self.responses)


def test_model_only_prepares_plan_and_human_execution_creates_output(tmp_path):
    source = tmp_path / "input.docx"
    doc = Document()
    doc.add_paragraph("The cost is 7")
    doc.save(source)
    original = source.read_bytes()
    editor = Editor(tmp_path)
    model = ScriptedModel(source)
    plans = []

    def dispatch(_name, arguments):
        plan = editor.plan(WordReplace.model_validate(arguments))
        plans.append(plan)
        return {"plan_id": plan.id, "preview": plan.public()}

    spec = ToolSpec(name="plan_word", description="Plan replacement without writing",
                    schema_model=WordReplace, access="write", allowed_paths=(source,))
    agent = OfficeAgent(model, dispatch, [spec], tmp_path / "state",
                        token_estimator=lambda _: 100)
    state = agent.run(agent.new_state("Replace cost with price", file_versions={str(source): digest(source)}))
    assert state.status == "done" and len(plans) == 1
    assert source.read_bytes() == original
    assert not list(tmp_path.glob("*.agent-*.docx"))
    result = editor.execute(plans[0].id, confirmed=True)
    assert result["status"] == "verified"
    assert Document(result["output"]).paragraphs[0].text == "The price is 7"
    assert source.read_bytes() == original
    assert editor.rollback(plans[0].id, confirmed=True)["status"] == "rolled_back"
