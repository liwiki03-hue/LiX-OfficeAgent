"""P3 planner tool wiring: fake service only; no provider call, network, .env or Office."""
import pytest
from docx import Document
from pydantic import ValidationError

from office_agent.agent import OfficeAgent, ToolCall
from office_agent.app import (ImageGenArgs, build_dispatcher, build_tools, select_tools,
                              split_read_write)
from office_agent.imagegen import ImageGenError, ImageRequest
from office_agent.writes import Editor

PROMPT = "synthetic private prompt blue squares"


@pytest.fixture
def source(tmp_path):
    """A real minimal DOCX: the tool stays behind the same readability gate as other tools."""
    path = tmp_path / "source.docx"
    Document().save(path)
    return path


class FakeService:
    """Records plan requests; refuses any attempt to reach a provider during dispatch."""

    def __init__(self, error=None):
        self.plans = []
        self.confirmations = []
        self._error = error

    def plan(self, request):
        if self._error is not None:
            raise self._error
        self.plans.append(request)
        return {"plan_id": "synthetic-plan-id", "kind": "image_generation",
                "effective_prompt": request.effective_prompt, "size": request.size,
                "style": request.style, "n": request.n, "confirmation_required": True,
                "save_confirmation_required": True, "cost_unknown": True,
                "provider_fees_irreversible": True}

    def confirm_generation(self, *args, **kwargs):
        self.confirmations.append((args, kwargs))
        raise AssertionError("dispatch must never authorize a provider call")


def test_schema_mirrors_image_request_and_rejects_everything_else(source):
    assert ImageGenArgs(path=str(source), prompt=PROMPT).n == 1
    assert ImageGenArgs(path=str(source), prompt=PROMPT, size="1024x1536", style="vivid").style == "vivid"
    for bad in [{"prompt": ""}, {"prompt": "   "}, {"prompt": "\ud800"}, {"prompt": "x" * 4001},
                {"prompt": PROMPT, "size": "512x512"}, {"prompt": PROMPT, "style": "bold"},
                {"prompt": PROMPT, "n": 2}, {"prompt": PROMPT, "n": True}, {"prompt": PROMPT, "n": "1"},
                {"prompt": PROMPT, "quality": "hd"}, {"prompt": PROMPT, "model": "x"}, {}]:
        with pytest.raises(ValidationError):
            ImageGenArgs(path=str(source), **bad)
    request = ImageRequest(prompt=PROMPT, style="vivid")
    assert ImageGenArgs(path=str(source), prompt=PROMPT, style="vivid").prompt == PROMPT
    assert request.effective_prompt == "Requested visual style: vivid.\n" + PROMPT


def test_catalogue_is_planner_only_and_never_in_read_tiers(source):
    tools = build_tools((source,))
    spec = next(tool for tool in tools if tool.name == "image_gen")
    assert spec.access == "write" and spec.schema_model is ImageGenArgs
    for kind in ("word", "spreadsheet"):
        tiered = select_tools(tools, kind)
        assert spec in tiered
        read_tools, write_tools = split_read_write(tiered)
        assert spec in write_tools and spec not in read_tools
        assert "image_gen" not in {tool.name for tool in read_tools}
    assert spec not in select_tools(tools, "legacy")
    assert "image_gen" not in {tool.name for tool in select_tools(tools, "unknown")}


def test_model_image_tool_omits_host_fixed_count_but_keeps_strict_schema(source):
    spec = next(tool for tool in build_tools((source,)) if tool.name == "image_gen")
    params = spec.model_schema()["function"]["parameters"]
    assert "n" not in params["properties"]
    assert "n" not in params.get("required", [])
    assert ImageGenArgs(path=str(source), prompt=PROMPT).n == 1
    with pytest.raises(ValidationError):
        ImageGenArgs(path=str(source), prompt=PROMPT, n="1")


def test_model_string_one_is_canonicalized_before_plan_not_at_dispatcher(source, tmp_path):
    service = FakeService()
    dispatch = build_dispatcher(Editor(source.parent), source, image_service_factory=lambda: service)

    class Model:
        def __init__(self):
            self.replies = iter([{"tool_calls": [{"name": "image_gen", "id": "image-1",
                "args": {"path": str(source), "prompt": PROMPT, "n": "1"}}]}, {"content": "done"}])

        def bind_tools(self, schemas):
            image_schema = next(s for s in schemas if s["function"]["name"] == "image_gen")
            assert "n" not in image_schema["function"]["parameters"]["properties"]
            return self

        def invoke(self, messages):
            return next(self.replies)

    spec = next(tool for tool in build_tools((source,)) if tool.name == "image_gen")
    agent = OfficeAgent(Model(), dispatch, [spec], tmp_path / "state", persist=False,
                        token_estimator=lambda _: 1)
    result = agent.run(agent.new_state("Generate one image"))
    assert result.status == "done" and not result.required.errors
    assert len(service.plans) == 1 and service.plans[0].n == 1
    assert not service.confirmations and not list(source.parent.glob("*.png"))
    # Repeated validation and dispatcher inputs stay integer-only, with no extra provider call.
    for bad in ("01", "1.0", "2", 1.0, True, None, {"value": 1}):
        with pytest.raises(ValidationError):
            agent._validate_plan(ToolCall(name="image_gen", args={
                "path": str(source), "prompt": PROMPT, "n": bad}), result)
        with pytest.raises(ValidationError):
            dispatch("image_gen", {"path": str(source), "prompt": PROMPT, "n": bad})
    assert len(service.plans) == 1


def test_dispatch_only_prepares_a_plan(source):
    service = FakeService()
    events = []
    dispatch = build_dispatcher(Editor(source.parent), source, on_plan=events.append,
                                image_service_factory=lambda: service)
    before = source.read_bytes()
    result = dispatch("image_gen", {"path": str(source), "prompt": PROMPT, "size": "1536x1024",
                                    "style": "vivid", "n": 1})
    assert result["plan_id"] == "synthetic-plan-id" and result["preview"]["plan_id"] == "synthetic-plan-id"
    assert "独立确认" in result["note"] and "未调用任何生图服务" in result["note"]
    assert service.plans[0].effective_prompt == "Requested visual style: vivid.\n" + PROMPT
    assert service.plans[0].size == "1536x1024" and service.confirmations == []
    assert events == [("image_plan", result["preview"])]
    assert source.read_bytes() == before
    assert not (source.parent / ".office_agent").exists()
    assert not list(source.parent.glob("*.png"))
    # The service is created once and reused, so its plan registry stays authoritative.
    assert dispatch("image_gen", {"path": str(source), "prompt": PROMPT})["plan_id"] == "synthetic-plan-id"
    assert len(service.plans) == 2


def test_unconfigured_or_broken_service_degrades_without_calling_anything(source):
    dispatch = build_dispatcher(Editor(source.parent), source)
    with pytest.raises(ValueError, match="云端生图未配置"):
        dispatch("image_gen", {"path": str(source), "prompt": PROMPT})
    def broken():
        raise ImageGenError("image_configuration_missing")
    with pytest.raises(ValueError, match="image_configuration_missing"):
        build_dispatcher(Editor(source.parent), source, image_service_factory=broken)(
            "image_gen", {"path": str(source), "prompt": PROMPT})
    def missing():
        raise ImageGenError("image_session_invalid")
    with pytest.raises(ValueError, match="image_session_invalid"):
        build_dispatcher(Editor(source.parent), source, image_service_factory=missing)(
            "image_gen", {"path": str(source), "prompt": PROMPT})
    assert not (source.parent / ".office_agent").exists()


@pytest.mark.parametrize("code,expected", [
    ("image_configuration_missing", "IMAGE_API_KEY"),
    ("image_configuration_invalid", "IMAGE_PROVIDER"),
])
def test_plan_time_configuration_errors_are_sanitized(source, code, expected):
    dispatch = build_dispatcher(Editor(source.parent), source,
                                image_service_factory=lambda: FakeService(error=ImageGenError(code)))
    with pytest.raises(ValueError) as error:
        dispatch("image_gen", {"path": str(source), "prompt": PROMPT})
    assert code in str(error.value) and expected in str(error.value)
    assert "synthetic-private-value" not in str(error.value)


def test_other_plan_errors_keep_the_stable_code(source):
    dispatch = build_dispatcher(Editor(source.parent), source,
                                image_service_factory=lambda: FakeService(error=ImageGenError("image_plan_limit")))
    with pytest.raises(ValueError, match="image_plan_limit"):
        dispatch("image_gen", {"path": str(source), "prompt": PROMPT})


def test_dispatch_applies_defaults_and_still_rejects_extra_or_missing_fields(source):
    service = FakeService()
    dispatch = build_dispatcher(Editor(source.parent), source, image_service_factory=lambda: service)
    result = dispatch("image_gen", {"path": str(source), "prompt": PROMPT})
    assert service.plans[-1].size == "1024x1024" and service.plans[-1].style == "natural"
    assert result["preview"]["n"] == 1
    with pytest.raises(ValidationError):
        dispatch("image_gen", {"path": str(source), "prompt": PROMPT, "quality": "hd"})
    with pytest.raises(ValidationError):
        dispatch("image_gen", {"path": str(source)})
    assert len(service.plans) == 1 and service.confirmations == []


def test_dispatch_requires_the_selected_readable_file(source, tmp_path):
    other = tmp_path / "other.docx"
    other.write_bytes(b"another synthetic file")
    dispatch = build_dispatcher(Editor(source.parent), source, image_service_factory=lambda: FakeService())
    with pytest.raises(ValueError, match="explicitly selected file"):
        dispatch("image_gen", {"path": str(other), "prompt": PROMPT})
    # Unreadable/risky files fail the same shared gate as every other tool.
    corrupt = tmp_path / "corrupt.docx"
    corrupt.write_bytes(b"not an Office container")
    broken = build_dispatcher(Editor(tmp_path), corrupt, image_service_factory=lambda: FakeService())
    with pytest.raises(ValueError, match="unreadable or risky"):
        broken("image_gen", {"path": str(corrupt), "prompt": PROMPT})
    assert not (source.parent / ".office_agent").exists()
