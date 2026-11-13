"""Day-2 compatibility matrix: synthetic bytes, fake transport, no paid requests."""
import base64
import io
import json
import threading

import httpx2
import pytest
from PIL import Image

from office_agent.imagegen import ImageGenConfig, ImageGenError, ImageRequest
from office_agent.imageprovider import OpenAIImageGenProvider
from office_agent import imageprovider
from office_agent.imagetransactions import ImageGenerationService


SESSION = "day2-session"
KEY = "synthetic-day2-only-key"


def pixels():
    buffer = io.BytesIO()
    with Image.new("RGB", (12, 8), "green") as image:
        image.save(buffer, format="PNG")
    return buffer.getvalue()


def inline(raw):
    return json.dumps({"data": [{"b64_json": base64.b64encode(raw).decode("ascii")}] }).encode()


@pytest.fixture
def context(tmp_path):
    source = tmp_path / "source.docx"
    source.write_bytes(b"synthetic source, never opened by Office")
    config = ImageGenConfig("openai_compatible", KEY, "https://images.example.test/v1", "synthetic-model")
    requests = []
    response = [lambda request: httpx2.Response(200, content=inline(pixels()))]

    def handle(request):
        requests.append(request)
        return response[0](request)

    def service(loader=lambda: config):
        return ImageGenerationService(tmp_path, source, SESSION, config_loader=loader,
            provider_factory=lambda: OpenAIImageGenProvider(
                transport_factory=lambda: httpx2.MockTransport(handle)))

    return tmp_path, source, requests, response, service


def approve(service):
    planned = service.plan(ImageRequest(prompt="synthetic green illustration"))
    assert planned["confirmation_required"] and planned["save_confirmation_required"]
    assert planned["cost_unknown"] and planned["provider_fees_irreversible"]
    return planned["plan_id"]


def assert_failure(service, plan_id, code, stage, detail, request_state):
    with pytest.raises(ImageGenError) as caught:
        service.confirm_generation(plan_id, SESSION, True)
    error = caught.value
    assert (error.code, error.stage, error.detail_code, error.request_state) == (
        code, stage, detail, request_state)
    with pytest.raises(ImageGenError, match="image_approval_consumed"):
        service.confirm_generation(plan_id, SESSION, True)
    with pytest.raises(ImageGenError, match="image_preview_unavailable"):
        service.preview(plan_id, SESSION)


def test_missing_configuration_fails_before_request_without_audit(context):
    root, source, requests, response, factory = context
    def missing():
        raise ImageGenError("image_configuration_missing")
    service = factory(loader=missing)
    with pytest.raises(ImageGenError, match="image_configuration_missing"):
        service.plan(ImageRequest(prompt="test"))
    assert requests == [] and not (root / ".office_agent").exists()
    assert source.read_bytes() == b"synthetic source, never opened by Office"


@pytest.mark.parametrize("case,expected_stage,expected_detail", [
    ("401", "provider_request", "image_provider_failed"),
    ("403", "provider_request", "image_provider_failed"),
    ("429", "provider_request", "image_provider_failed"),
    ("500", "provider_request", "image_provider_failed"),
    ("disconnect", "provider_request", "image_provider_failed"),
    ("timeout", "response", "image_response_timeout"),
    ("encoding", "response", "image_response_encoding"),
    ("oversize", "response", "image_response_too_large"),
    ("url", "response", "image_response_invalid"),
    ("base64", "response", "image_response_invalid"),
    ("duplicate", "response", "image_response_invalid"),
    ("multiple", "response", "image_response_invalid"),
    ("dimensions", "validation", "image_raster_invalid"),
])
def test_one_attempt_failure_matrix(context, monkeypatch, case, expected_stage, expected_detail, caplog):
    root, source, requests, response, factory = context
    original = source.read_bytes()
    service = factory()
    plan_id = approve(service)
    if case in {"401", "403", "429", "500"}:
        response[0] = lambda request: httpx2.Response(
            int(case), headers={"x-request-id": KEY, "retry-after": "1"},
            content=(KEY + " provider secret").encode())
    elif case == "disconnect":
        response[0] = lambda request: (_ for _ in ()).throw(httpx2.ConnectError(KEY))
    elif case == "timeout":
        response[0] = lambda request: (_ for _ in ()).throw(httpx2.ReadTimeout(KEY))
    elif case == "encoding":
        response[0] = lambda request: httpx2.Response(
            200, headers={"content-encoding": "gzip"}, stream=httpx2.ByteStream(inline(pixels())))
    elif case == "oversize":
        response[0] = lambda request: httpx2.Response(
            200, headers={"content-length": str(imageprovider.MAX_RESPONSE_BYTES + 1)},
            content=inline(pixels()))
    elif case == "url":
        response[0] = lambda request: httpx2.Response(
            200, content=b'{"data":[{"url":"https://images.example.test/private"}]}')
    elif case == "base64":
        response[0] = lambda request: httpx2.Response(200, content=b'{"data":[{"b64_json":"bad!"}]}')
    elif case == "duplicate":
        response[0] = lambda request: httpx2.Response(
            200, content=b'{"data":[{"b64_json":"a","b64_json":"b"}]}')
    elif case == "multiple":
        response[0] = lambda request: httpx2.Response(
            200, content=json.dumps({"data": [{"b64_json": base64.b64encode(pixels()).decode()}] * 2}).encode())
    else:
        monkeypatch.setattr(imageprovider, "MAX_DIMENSION", 11)
    caplog.set_level("DEBUG")
    assert_failure(service, plan_id, "image_generation_failed", expected_stage,
                   expected_detail, "attempt_committed")
    assert len(requests) == 1 and requests[0].method == "POST"
    assert requests[0].url.path.endswith("/images/generations")
    assert json.loads(requests[0].content)["n"] == 1
    audit = json.loads((root / ".office_agent" / "imagegen" / plan_id / "audit.json").read_text(encoding="utf-8"))
    assert audit["status"] == "failed" and audit["provider_request_state"] == "attempt_committed"
    # Dispatch may have reached and billed the remote provider; HTTP failure is not a refund receipt.
    assert audit["cost_unknown"] is True and audit["provider_fees_irreversible"] is True
    assert audit["saved"] is False and "billed" not in audit
    assert KEY not in json.dumps(audit) and KEY not in caplog.text
    assert original == source.read_bytes() and not list(root.glob("*.png"))


def test_tampered_audit_discards_successful_reply_without_save(context):
    root, source, requests, response, factory = context
    service = factory()
    plan_id = approve(service)
    audit_path = root / ".office_agent" / "imagegen" / plan_id / "audit.json"
    def tamper(request):
        audit_path.write_text('{"foreign":true}', encoding="utf-8")
        return httpx2.Response(200, content=inline(pixels()))
    response[0] = tamper
    assert_failure(service, plan_id, "image_audit_changed", "provider_request", None,
                   "attempt_committed")
    assert audit_path.read_text(encoding="utf-8") == '{"foreign":true}'
    assert len(requests) == 1 and not list(root.glob("*.png"))


def test_session_invalidated_during_request_discards_late_reply(context):
    root, source, requests, response, factory = context
    service = factory()
    plan_id = approve(service)
    entered, release = threading.Event(), threading.Event()
    results = []
    def hold(request):
        entered.set()
        if not release.wait(5):
            raise AssertionError("fake response was not released")
        return httpx2.Response(200, content=inline(pixels()))
    response[0] = hold
    def work():
        try:
            service.confirm_generation(plan_id, SESSION, True)
        except ImageGenError as error:
            results.append((error.code, error.request_state))
    worker = threading.Thread(target=work, daemon=True)
    worker.start()
    try:
        assert entered.wait(5)
        service.invalidate()
    finally:
        release.set()
        worker.join(6)
    assert not worker.is_alive()
    assert results == [("image_operation_invalidated", "attempt_committed")]
    assert len(requests) == 1 and not list(root.glob("*.png"))
    assert source.read_bytes() == b"synthetic source, never opened by Office"
    with pytest.raises(ImageGenError, match="image_operation_invalidated"):
        service.plan_save(plan_id, SESSION)


def test_success_is_unsaved_until_separate_save_approval(context):
    root, source, requests, response, factory = context
    service = factory()
    plan_id = approve(service)
    original = source.read_bytes()
    assert service.confirm_generation(plan_id, SESSION, True)["status"] == "generated"
    assert len(requests) == 1 and not list(root.glob("*.png"))
    cancelled = service.plan_save(plan_id, SESSION)
    assert service.confirm_save(cancelled["save_id"], SESSION, False)["status"] == "cancelled"
    assert not list(root.glob("*.png"))
    fresh = service.plan_save(plan_id, SESSION)
    saved = service.confirm_save(fresh["save_id"], SESSION, True)
    with Image.open(saved["path"]) as image:
        image.load()
        assert image.format == "PNG" and image.size == (12, 8)
    assert len(requests) == 1 and source.read_bytes() == original
