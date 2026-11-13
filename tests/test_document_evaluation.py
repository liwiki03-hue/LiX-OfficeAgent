from copy import deepcopy

import pytest

from examples.eval_documents import CASES, ScriptedModel, run_case, run_suite
from examples.eval_documents import make_fixture
from office_agent.app import build_dispatcher
from office_agent.writes import Editor, GuardError, WordReplace, digest


@pytest.mark.parametrize("case", CASES, ids=lambda case: case.id)
def test_document_golden_case(case, tmp_path):
    report = run_case(case, tmp_path / case.id)
    assert report["passed"], report
    assert report["scores"]["source_unchanged"] is True
    if case.tool.startswith("plan_") and not case.rejected:
        assert report["output_verified"] is True
        assert report["rollback_verified"] is True
        assert report["scores"]["fail_closed"] is True


def test_golden_dimensions_and_verification_labels(tmp_path):
    report = run_suite(tmp_path)
    assert report["case_count"] == 26
    assert report["passed"] is True
    assert report["mode"] == "scripted_model_real_documents"
    assert report["provider_accuracy"] is None
    assert report["human_confirmation_verified"] is False
    assert report["dimensions"]["answer_correct"]["evaluated"] == 26
    assert report["dimensions"]["plan_parameters_correct"]["evaluated"] == 14
    assert report["dimensions"]["fail_closed"]["evaluated"] == 18
    assert all(value["rate"] == 1 for value in report["dimensions"].values())


def test_wrong_model_answer_is_not_counted_as_correct(tmp_path):
    case = CASES[0]

    def factory(case, source):
        model = ScriptedModel(case, {"path": str(source), **case.args})
        messages = list(model.messages)
        messages[-1] = {"content": "预算9999"}
        model.messages = iter(messages)
        return model

    report = run_case(case, tmp_path / case.id, model_factory=factory)
    assert report["passed"] is False
    assert report["scores"]["answer_correct"] is False
    assert report["read_verified"] is True


def test_wrong_plan_is_not_executed_even_with_correct_final_answer(tmp_path):
    case = next(case for case in CASES if case.id == "excel_set")

    def factory(case, source):
        args = {"path": str(source), **deepcopy(case.args), "value": 9999}
        return ScriptedModel(case, args)

    report = run_case(case, tmp_path / case.id, model_factory=factory)
    assert report["passed"] is False
    assert report["scores"]["plan_parameters_correct"] is False
    assert report["output_verified"] is None
    assert not list(tmp_path.rglob("*.agent-*"))
    assert report["scores"]["source_unchanged"] is True


@pytest.mark.parametrize("approval", [False, None, 0, 1, "true", []])
def test_production_dispatcher_enforces_strict_confirmation_for_simple_edits(tmp_path, approval):
    source = make_fixture(tmp_path, "word")
    original = digest(source)
    editor = Editor(tmp_path)
    build_dispatcher(editor, source)
    plan = editor.plan(WordReplace(path=str(source), old="预算", new="金额", paragraph_index=0))
    assert plan.confirmation_required is True
    with pytest.raises(GuardError, match="human confirmation"):
        editor.execute(plan.id, confirmed=approval)
    assert digest(source) == original
    assert not editor.audit.exists()
    assert plan.id in editor.plans
