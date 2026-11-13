"""Mock transport + synthetic pixels only; no network or real credential access."""
import base64
import io
import json

import httpx2
import pytest
from PIL import Image, PngImagePlugin

from office_agent.imagegen import ImageGenConfig, ImageGenError, ImageRequest
from office_agent import imageprovider as provider

KEY = "synthetic-provider-only-key"


def config(response_format="auto"):
    return ImageGenConfig("openai_compatible", KEY, "https://images.example.test/v1", "synthetic-model",
                          response_format=response_format)


def raster(fmt="PNG", metadata=False):
    output = io.BytesIO()
    with Image.new("RGB", (12, 8), "blue") as image:
        options = {}
        if metadata and fmt == "PNG":
            text = PngImagePlugin.PngInfo()
            text.add_text("private", "DO-NOT-PERSIST")
            options["pnginfo"] = text
        image.save(output, format=fmt, **options)
    return output.getvalue()


def response_body(raw=None):
    return json.dumps({"data": [{"b64_json": base64.b64encode(raw or raster()).decode()}]}).encode()


@pytest.mark.parametrize("fmt", ["PNG", "JPEG"])
def test_normalization_strips_metadata_and_reopens(fmt):
    result = provider.normalize_generated_image(raster(fmt, metadata=True))
    assert result.metadata["width"] == 12 and result.metadata["height"] == 8
    assert result.metadata["metadata_stripped"] is True
    assert b"DO-NOT-PERSIST" not in result.raw
    with Image.open(io.BytesIO(result.raw)) as image:
        image.load()
        assert image.format == "PNG" and not image.info


@pytest.mark.parametrize("raw", [b"", b"not an image", b'<svg onload="alert(1)"/>', b"GIF89a", b"\x89PNG\r\n\x1a\n"])
def test_bad_rasters_rejected(raw):
    with pytest.raises(ImageGenError):
        provider.normalize_generated_image(raw)


def test_resource_caps_and_animation(monkeypatch):
    raw = raster()
    monkeypatch.setattr(provider, "MAX_PIXELS", 95)
    with pytest.raises(ImageGenError):
        provider.normalize_generated_image(raw)
    monkeypatch.setattr(provider, "MAX_PIXELS", 40_000_000)
    monkeypatch.setattr(provider, "MAX_DIMENSION", 11)
    with pytest.raises(ImageGenError):
        provider.normalize_generated_image(raw)
    monkeypatch.setattr(provider, "MAX_DIMENSION", 32768)
    monkeypatch.setattr(provider, "MAX_IMAGE_BYTES", 5)
    with pytest.raises(ImageGenError):
        provider.normalize_generated_image(raw)
    monkeypatch.setattr(provider, "MAX_IMAGE_BYTES", 10_000_000)
    output = io.BytesIO()
    with Image.new("RGB", (12, 8), "blue") as first, Image.new("RGB", (12, 8), "red") as second:
        first.save(output, format="PNG", save_all=True, append_images=[second])
    with pytest.raises(ImageGenError):
        provider.normalize_generated_image(output.getvalue())


@pytest.mark.parametrize("body", [
    b"null", b"[]", b"{}", b'{"data":[]}', b'{"data":[{},{}]}',
    b'{"data":[{"url":"https://private.example.test/"}]}',
    b'{"data":[{"b64_json":"YQ==","url":null}]}', b'{"data":[{"b64_json":"YQ= "}]}',
    b'{"data":[{"b64_json":"YR=="}]}', b'{"data":[{"b64_json":""}]}',
    b'{"data":[{"b64_json":123}]}', b'{"data":[],"data":[{"b64_json":"YQ=="}]}',
    b"\xff", b'{"data":[{"b64_json":"data:image/png;base64,YQ=="}]}',
])
def test_bad_response_rejected_without_url_fetch(body):
    with pytest.raises(ImageGenError) as error:
        provider.decode_image_response(body)
    assert str(error.value) == "image_response_invalid"


def test_inline_decoding_is_exact_and_bounded(monkeypatch):
    raw = raster()
    assert provider.decode_image_response(response_body(raw)) == raw
    monkeypatch.setattr(provider, "MAX_IMAGE_BYTES", 1)
    with pytest.raises(ImageGenError):
        provider.decode_image_response(response_body(raw))


@pytest.mark.parametrize("format", ["auto", "b64_json"])
def test_sdk_single_request_parameters_and_clean_result(format, monkeypatch):
    monkeypatch.setenv("OPENAI_API_KEY", "synthetic-private-chat")
    monkeypatch.setenv("OPENAI_ADMIN_KEY", "synthetic-private-admin")
    monkeypatch.setenv("OPENAI_ORG_ID", "synthetic-private-org")
    monkeypatch.setenv("OPENAI_PROJECT_ID", "synthetic-private-project")
    monkeypatch.setenv("OPENAI_WEBHOOK_SECRET", "synthetic-private-webhook")
    monkeypatch.setenv("OPENAI_CUSTOM_HEADERS", "Host: attacker.test\nAuthorization: Bearer synthetic-private-chat\nX-Chat-Key: synthetic-private-chat\nContent-Length: 9")
    calls = []
    def handle(request):
        calls.append(request)
        assert request.method == "POST"
        assert str(request.url) == "https://images.example.test/v1/images/generations"
        assert request.headers["authorization"] == "Bearer " + KEY
        assert request.headers["accept-encoding"] == "identity"
        assert request.headers["host"] == "images.example.test"
        assert int(request.headers["content-length"]) == len(request.content)
        assert "synthetic-private" not in json.dumps(dict(request.headers))
        assert "openai-organization" not in request.headers and "openai-project" not in request.headers
        body = json.loads(request.content)
        assert body["n"] == 1 and body["model"] == "synthetic-model"
        assert body["size"] == "1024x1024"
        assert body["prompt"] == ImageRequest(prompt="blue").effective_prompt
        assert "style" not in body
        if format == "auto":
            assert "response_format" not in body
        else:
            assert body["response_format"] == "b64_json"
        return httpx2.Response(200, content=response_body(raster(metadata=True)))
    result = provider.OpenAIImageGenProvider(transport_factory=lambda: httpx2.MockTransport(handle)).generate(
        ImageRequest(prompt="blue"), config(format))
    assert len(calls) == 1 and b"DO-NOT-PERSIST" not in result
    provider.normalize_generated_image(result)


@pytest.mark.parametrize("status", [301, 429, 500])
def test_no_redirect_or_retry_and_sanitized_error(status, caplog):
    caplog.set_level("DEBUG")
    calls = []
    def handle(request):
        calls.append(request)
        return httpx2.Response(status, headers={"location": "https://private.example.test/", "x-request-id": KEY},
                               content=(KEY + " malicious-provider-error").encode())
    with pytest.raises(ImageGenError) as error:
        provider.OpenAIImageGenProvider(transport_factory=lambda: httpx2.MockTransport(handle)).generate(
            ImageRequest(prompt="blue"), config())
    assert len(calls) == 1
    assert str(error.value) == "image_provider_failed" and KEY not in repr(error.value)
    assert KEY not in caplog.text and "malicious-provider-error" not in caplog.text


class Chunks(httpx2.SyncByteStream):
    def __init__(self, values):
        self.values, self.closed = values, False
    def __iter__(self):
        yield from self.values
    def close(self):
        self.closed = True


@pytest.mark.parametrize("status", [200, 500])
@pytest.mark.parametrize("headers,chunks", [
    ({}, [b"a" * 40, b"b" * 40]), ({"content-length": "1000"}, [b"a"]),
    ({"content-length": "invalid"}, [b"a"]), ({"content-encoding": "gzip"}, [b"a"]),
])
def test_transport_bounds_even_error_bodies_and_closes(monkeypatch, status, headers, chunks):
    monkeypatch.setattr(provider, "MAX_RESPONSE_BYTES", 64)
    stream = Chunks(chunks)
    calls = []
    def handle(request):
        calls.append(request)
        return httpx2.Response(status, headers=headers, stream=stream)
    with pytest.raises(ImageGenError) as error:
        provider.OpenAIImageGenProvider(transport_factory=lambda: httpx2.MockTransport(handle)).generate(
            ImageRequest(prompt="blue"), config())
    assert str(error.value) in {"image_response_too_large", "image_response_encoding"}
    assert len(calls) == 1 and stream.closed


def test_transport_deadline_uses_monotonic_and_closes(monkeypatch):
    ticks = iter([0, 61])
    monkeypatch.setattr(provider.time, "monotonic", lambda: next(ticks))
    stream = Chunks([b"data"])
    transport = provider._BoundedTransport(httpx2.MockTransport(
        lambda request: httpx2.Response(200, stream=stream)), 60)
    try:
        with pytest.raises(ImageGenError, match="image_response_timeout"):
            transport.handle_request(httpx2.Request("POST", "https://images.example.test"))
        assert stream.closed
    finally:
        transport.close()


def test_output_writer_cap_before_allocation(monkeypatch):
    monkeypatch.setattr(provider, "MAX_IMAGE_BYTES", 3)
    buffer = provider._BoundedImageBuffer()
    with pytest.raises(ImageGenError, match="image_raster_too_large"):
        buffer.write(b"four")
    assert buffer.getvalue() == b""


def test_default_transport_does_not_inherit_environment_tls_or_proxy(monkeypatch):
    options = []
    def fake_transport(**kwargs):
        options.append(kwargs)
        return httpx2.MockTransport(lambda request: httpx2.Response(200, content=b"{}"))
    monkeypatch.setattr(provider.httpx2, "HTTPTransport", fake_transport)
    instance = provider.OpenAIImageGenProvider()
    transport = instance._transport_factory()
    try:
        assert options == [{"retries": 0, "trust_env": False}]
    finally:
        transport.close()
