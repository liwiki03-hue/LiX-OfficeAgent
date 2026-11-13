"""Host-only P3 provider primitives; not registered and not an approval endpoint.

A future transaction service must authorize/audit BEFORE invoking generate().
Timeout is per-I/O, with an additional monotonic response deadline checked at
headers/chunk boundaries; it is not a hard process-level wall-clock deadline.
"""
from __future__ import annotations

import base64
import hashlib
import io
import json
import time
import warnings

import httpx2
from openai import OpenAI

from .imagegen import GeneratedImage, ImageGenConfig, ImageGenError, ImageRequest

MAX_RESPONSE_BYTES = 20_000_000
MAX_IMAGE_BYTES = 10_000_000
MAX_PIXELS = 40_000_000
MAX_DIMENSION = 32768


class _BoundedTransport(httpx2.BaseTransport):
    """Cap even error bodies BEFORE the SDK allocates/parses/logs their contents."""
    def __init__(self, inner: httpx2.BaseTransport, timeout: float, api_key: str | None = None):
        self.inner, self.timeout, self._api_key = inner, timeout, api_key

    def handle_request(self, request):
        if self._api_key is not None:
            # SDK ambient/custom headers must never leak unrelated chat credentials.
            request.headers.clear()
            request.headers.update({"Host": request.url.netloc.decode("ascii"),
                                    "Authorization": "Bearer " + self._api_key,
                                    "Content-Type": "application/json", "Accept": "application/json",
                                    "Accept-Encoding": "identity", "User-Agent": "OfficeAgentImageGen/1",
                                    "Content-Length": str(len(request.content)),
                                    "X-Stainless-Raw-Response": "stream"})
        deadline = time.monotonic() + self.timeout
        response = self.inner.handle_request(request)
        try:
            if time.monotonic() > deadline:
                raise ImageGenError("image_response_timeout")
            if response.headers.get("content-encoding", "identity").lower().strip() != "identity":
                raise ImageGenError("image_response_encoding")
            length = response.headers.get("content-length")
            if length is not None and (not length.isdecimal() or int(length) > MAX_RESPONSE_BYTES):
                raise ImageGenError("image_response_too_large")
            body = bytearray()
            for chunk in response.stream:
                if time.monotonic() > deadline:
                    raise ImageGenError("image_response_timeout")
                if len(body) + len(chunk) > MAX_RESPONSE_BYTES:
                    raise ImageGenError("image_response_too_large")
                body.extend(chunk)
            if time.monotonic() > deadline:
                raise ImageGenError("image_response_timeout")
            # Do not let provider-controlled request IDs/headers reach SDK logs.
            safe_headers = {"Content-Type": "application/json", "Content-Length": str(len(body))}
            return httpx2.Response(response.status_code, headers=safe_headers,
                                   content=bytes(body), request=request)
        finally:
            response.close()

    def close(self):
        self.inner.close()


def _unique_object(pairs):
    value = {}
    for key, item in pairs:
        if key in value:
            raise ValueError("duplicate_json_key")
        value[key] = item
    return value


def decode_image_response(body: bytes) -> bytes:
    """Exactly one canonical inline base64 image; never fetch provider URLs."""
    try:
        if type(body) is not bytes or len(body) > MAX_RESPONSE_BYTES:
            raise ValueError()
        value = json.loads(body.decode("utf-8"), object_pairs_hook=_unique_object)
        data = value["data"]
        if type(data) is not list or len(data) != 1 or type(data[0]) is not dict:
            raise ValueError()
        entry = data[0]
        if "url" in entry:
            raise ValueError()
        text = entry["b64_json"]
        if type(text) is not str or not text or len(text) > ((MAX_IMAGE_BYTES + 2) // 3) * 4:
            raise ValueError()
        raw = base64.b64decode(text, validate=True)
        if not raw or len(raw) > MAX_IMAGE_BYTES or base64.b64encode(raw).decode("ascii") != text:
            raise ValueError()
        return raw
    except Exception:
        raise ImageGenError("image_response_invalid") from None


class _BoundedImageBuffer(io.BytesIO):
    def write(self, value):
        if self.tell() + len(value) > MAX_IMAGE_BYTES:
            raise ImageGenError("image_raster_too_large")
        return super().write(value)


def normalize_generated_image(raw: bytes) -> GeneratedImage:
    """Verify/reopen/load a static PNG/JPEG and encode a metadata-free PNG."""
    try:
        from PIL import Image
    except ImportError:
        raise ImageGenError("image_decoder_unavailable") from None
    if type(raw) is not bytes or not raw or len(raw) > MAX_IMAGE_BYTES:
        raise ImageGenError("image_raster_too_large")
    try:
        with warnings.catch_warnings():
            warnings.simplefilter("error", Image.DecompressionBombWarning)
            def check(image):
                width, height = image.size
                if (image.format not in {"PNG", "JPEG"} or getattr(image, "n_frames", 1) != 1
                        or not 0 < width <= MAX_DIMENSION or not 0 < height <= MAX_DIMENSION
                        or width * height > MAX_PIXELS):
                    raise ValueError()
            with Image.open(io.BytesIO(raw)) as image:
                check(image)
                image.verify()
            with Image.open(io.BytesIO(raw)) as image:
                check(image)
                image.load()
                mode = "RGBA" if "A" in image.getbands() or "transparency" in image.info else "RGB"
                with Image.new(mode, image.size) as clean:
                    with image.convert(mode) as converted:
                        clean.paste(converted)
                    buffer = _BoundedImageBuffer()
                    clean.save(buffer, format="PNG")
                    png = buffer.getvalue()
            with Image.open(io.BytesIO(png)) as reopened:
                check(reopened)
                reopened.load()
                width, height = reopened.size
                if reopened.info:
                    raise ValueError()
        return GeneratedImage(png, {"mime_type": "image/png", "format": "PNG", "width": width,
                                   "height": height, "bytes": len(png),
                                   "sha256": hashlib.sha256(png).hexdigest(), "metadata_stripped": True})
    except ImageGenError:
        raise
    except Exception:
        raise ImageGenError("image_raster_invalid") from None


class OpenAIImageGenProvider:
    """One request, no retries/redirects/proxies. Host must obtain approval first."""
    def __init__(self, *, transport_factory=None):
        self._transport_factory = transport_factory or (lambda: httpx2.HTTPTransport(retries=0, trust_env=False))

    def generate(self, request: ImageRequest, config: ImageGenConfig) -> bytes:
        if not isinstance(request, ImageRequest) or not isinstance(config, ImageGenConfig):
            raise ImageGenError("image_provider_invalid_arguments")
        try:
            with httpx2.Client(transport=_BoundedTransport(self._transport_factory(), config.timeout, config.api_key),
                               timeout=config.timeout, follow_redirects=False, trust_env=False,
                               headers={"Accept-Encoding": "identity"}) as http:
                with OpenAI(api_key=config.api_key, base_url=config.base_url, max_retries=0,
                            admin_api_key="", organization="", project="", webhook_secret="",
                            default_headers={"Authorization": "Bearer " + config.api_key},
                            timeout=config.timeout, http_client=http) as client:
                    kwargs = {"prompt": request.effective_prompt, "model": config.model,
                              "n": 1, "size": request.size}
                    if config.response_format == "b64_json":
                        kwargs["response_format"] = "b64_json"
                    with client.images.with_streaming_response.generate(**kwargs) as response:
                        body = bytearray()
                        for chunk in response.iter_bytes():
                            if len(body) + len(chunk) > MAX_RESPONSE_BYTES:
                                raise ImageGenError("image_response_too_large")
                            body.extend(chunk)
            return normalize_generated_image(decode_image_response(bytes(body))).raw
        except Exception as error:
            safe = {"image_response_timeout", "image_response_encoding", "image_response_too_large",
                    "image_response_invalid", "image_decoder_unavailable", "image_raster_too_large",
                    "image_raster_invalid"}
            current = error
            for _ in range(8):
                if isinstance(current, ImageGenError) and current.code in safe:
                    raise ImageGenError(current.code) from None
                if isinstance(current, httpx2.TimeoutException):
                    raise ImageGenError("image_response_timeout") from None
                current = current.__cause__
                if current is None:
                    break
            status = getattr(error, 'status_code', None)
            detail = ('image_authentication_failed' if status in (401, 403) else
                      'image_model_not_found' if status == 404 else
                      'image_rate_limited' if status == 429 else
                      'image_service_failed' if status and status >= 500 else
                      'image_address_unreachable' if 'connect' in type(error).__name__.lower() else None)
            raise ImageGenError("image_provider_failed", provider_code=detail) from None
