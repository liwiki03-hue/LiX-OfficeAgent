"""Natural-language image batches remain previews until per-image human consent."""
import io

import pytest
from PIL import Image
from docx import Document
from pydantic import ValidationError

from office_agent.app import build_dispatcher
from office_agent.clarification import classify_intent, image_request_count
from office_agent.imagegen import ImageGenConfig
from office_agent.imagetransactions import ImageGenerationService
from office_agent.writes import Editor


@pytest.mark.parametrize("utterance, expected", [
    ("生成一个大肥鱼的图片", 1),
    ("请画三张海边的照片", 3),
    ("为文档生成 4 张插画", 4),
    ("Generate 2 images of a blue fish", 2),
    ("Draw three pictures of a blue fish", 3),
    ("画一张猫、一张狗的照片", 2),
    ("生成一张 1024x1024 的鱼图片", 1),
    ("生成3D鱼图片", 1),
    ("画一张有三个人的照片", 1),
    ("请画两张，一张是猫，一张是狗", 2),
    ("生成图片", 1),
])
def test_count_comes_from_user_not_model(utterance, expected):
    assert image_request_count(utterance) == expected


@pytest.mark.parametrize("utterance", ["生成9张猫图片", "生成十张图片", "generate 999 images", "画几张鱼的照片",
                                   "draw several images", "生成2张猫图片和3张狗图片"])
def test_unsupported_or_unspecified_plural_does_not_silently_bill(utterance):
    with pytest.raises(ValueError, match="张数"):
        image_request_count(utterance)


def test_clear_subject_needs_no_extra_style_or_count_question():
    assert not classify_intent("生成一个大肥鱼的图片").needs_clarification
    assert not classify_intent("请画三张海边的照片").needs_clarification
    assert classify_intent("生成图片").needs_clarification


class RecordingService:
    def __init__(self):
        self.requests = []
        self.provider_calls = 0

    def plan(self, request):
        self.requests.append(request)
        return {"plan_id": f"plan-{len(self.requests)}", "effective_prompt": request.effective_prompt,
                "n": request.n, "size": request.size, "style": request.style,
                "confirmation_required": True, "save_confirmation_required": True,
                "cost_unknown": True}


def test_one_model_call_creates_three_distinct_unpaid_plans(tmp_path):
    source = tmp_path / "selected.docx"
    Document().save(source)
    service = RecordingService()
    events = []
    dispatch = build_dispatcher(Editor(tmp_path), source, image_count=3, on_plan=events.append,
                                image_service_factory=lambda: service)
    result = dispatch("image_gen", {"path": str(source), "prompt": "Three animal portraits",
                                   "prompts": ["A cat portrait", "A dog portrait", "A fish portrait"]})
    assert result["plan_id"] == "plan-1"
    assert result["plan_ids"] == ["plan-1", "plan-2", "plan-3"]
    assert [r.prompt for r in service.requests] == ["A cat portrait", "A dog portrait", "A fish portrait"]
    assert [r.n for r in service.requests] == [1, 1, 1]
    assert [event[1]["plan_id"] for event in events] == result["plan_ids"]
    assert all(event[1]["confirmation_required"] for event in events)
    assert service.provider_calls == 0 and not list(tmp_path.glob("*.png"))


def test_one_prompt_creates_bounded_independent_samples_and_rejects_mismatch(tmp_path):
    source = tmp_path / "selected.docx"
    Document().save(source)
    service = RecordingService()
    dispatch = build_dispatcher(Editor(tmp_path), source, image_count=2,
                                image_service_factory=lambda: service)
    for prompts in (["one"], ["one", "two", "three"]):
        with pytest.raises(ValueError, match="张数"):
            dispatch("image_gen", {"path": str(source), "prompt": "one subject", "prompts": prompts})
    with pytest.raises(ValidationError):
        dispatch("image_gen", {"path": str(source), "prompt": "one subject", "n": 2})
    assert not service.requests
    result = dispatch("image_gen", {"path": str(source), "prompt": "one subject"})
    assert result["plan_ids"] == ["plan-1", "plan-2"]
    assert [r.prompt for r in service.requests] == ["one subject", "one subject"]


def test_real_host_plans_require_distinct_fee_and_save_decisions(tmp_path):
    source = tmp_path / "selected.docx"
    Document().save(source)
    provider_calls = []

    class OfflineProvider:
        def generate(self, request, config):
            provider_calls.append(request)
            raw = io.BytesIO()
            Image.new("RGB", (8, 8), "blue").save(raw, format="PNG")
            return raw.getvalue()

    service = ImageGenerationService(
        tmp_path, source, "session", config_loader=lambda: ImageGenConfig(
            "openai_compatible", "synthetic-key", "https://image.example.test/v1", "synthetic-model"),
        provider_factory=OfflineProvider)
    events = []
    dispatch = build_dispatcher(Editor(tmp_path), source, image_count=3,
                                on_plan=events.append, image_service_factory=lambda: service)
    result = dispatch("image_gen", {"path": str(source), "prompt": "three different fish"})
    first, second, third = result["plan_ids"]
    assert len({first, second, third}) == 3 and len(events) == 3
    assert provider_calls == [] and not list(tmp_path.glob("*.png"))
    assert service.confirm_generation(first, "session", False)["status"] == "cancelled"
    assert provider_calls == []  # cancelling one cannot generate another
    assert service.confirm_generation(second, "session", True)["status"] == "generated"
    assert len(provider_calls) == 1
    save = service.plan_save(second, "session")
    assert not list(tmp_path.glob("*.png"))
    assert service.confirm_save(save["save_id"], "session", False)["status"] == "cancelled"
    assert not list(tmp_path.glob("*.png"))
    assert service.confirm_generation(third, "session", True)["status"] == "generated"
    assert len(provider_calls) == 2
    save = service.plan_save(third, "session")
    assert service.confirm_save(save["save_id"], "session", True)["status"] == "saved"
    assert len(list(tmp_path.glob("*.png"))) == 1
