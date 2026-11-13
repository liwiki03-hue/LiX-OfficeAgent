"""P3 foundation tests: injected configuration only, no network or credential files."""
import json

import pytest
from pydantic import ValidationError

from office_agent.imagegen import GeneratedImage, ImageGenConfig, ImageGenError, ImageRequest


KEY = "synthetic-only-never-a-real-key"


def environment(**changes):
    data = {"IMAGE_PROVIDER": "openai_compatible", "IMAGE_API_KEY": KEY,
            "IMAGE_BASE_URL": "https://images.example.test/v1", "IMAGE_MODEL": "synthetic-image"}
    data.update(changes)
    return data


def test_request_and_exact_effective_prompt():
    request = ImageRequest(prompt="画一张蓝色图片")
    assert request.n == 1
    assert request.effective_prompt == "Requested visual style: natural.\n画一张蓝色图片"
    assert ImageRequest(prompt="x" * 4000, size="1536x1024", style="vivid").prompt == "x" * 4000
    assert ImageRequest(prompt="😀" * 4000).prompt == "😀" * 4000


@pytest.mark.parametrize("changes", [
    {"prompt": ""}, {"prompt": " \n"}, {"prompt": "x" * 4001}, {"prompt": "\ud800"},
    {"prompt": 123}, {"n": True}, {"n": "1"}, {"n": 0}, {"n": 2},
    {"size": "512x512"}, {"style": "untrusted"}, {"provider": "arbitrary"},
    {"api_key": "forbidden"}, {"output_path": "escape.png"}, {"confirmed": True},
])
def test_strict_request_rejects(changes):
    with pytest.raises(ValidationError):
        ImageRequest(**({"prompt": "synthetic"} | changes))


def test_injected_config_never_loads_dotenv(monkeypatch):
    import dotenv
    def forbidden(*args, **kwargs):
        raise AssertionError("real dotenv must never be read")
    monkeypatch.setattr(dotenv, "load_dotenv", forbidden)
    config = ImageGenConfig.from_env(environment())
    assert config.timeout == 60
    assert config.response_format == "auto"
    assert KEY not in repr(config)
    assert KEY not in json.dumps(config.public_config())
    assert "api_key" not in config.public_config()
    assert "fingerprint" not in config.public_config()


@pytest.mark.parametrize("name", ["IMAGE_PROVIDER", "IMAGE_API_KEY", "IMAGE_BASE_URL", "IMAGE_MODEL"])
def test_required_config_missing_has_sanitized_error(name):
    data = environment()
    del data[name]
    with pytest.raises(ImageGenError) as caught:
        ImageGenConfig.from_env(data)
    assert str(caught.value) == "image_configuration_missing"
    assert KEY not in str(caught.value)


@pytest.mark.parametrize("changes", [
    {"IMAGE_PROVIDER": "unknown"}, {"IMAGE_BASE_URL": "http://images.example.test/v1"},
    {"IMAGE_BASE_URL": "https://user:password@images.example.test/v1"},
    {"IMAGE_BASE_URL": "https://images.example.test/v1?token=" + KEY},
    {"IMAGE_BASE_URL": "https://images.example.test/v1#fragment"},
    {"IMAGE_BASE_URL": "https://images.example.test/v1?"},
    {"IMAGE_BASE_URL": "https://images.example.test/v1#"},
    {"IMAGE_BASE_URL": "https://images.example.test:99999/v1"},
    {"IMAGE_BASE_URL": "https://images.example.test:0/v1"},
    {"IMAGE_BASE_URL": "https://images.example.test/\nunsafe"},
    {"IMAGE_BASE_URL": "https://images.example.test\\unsafe"},
    {"IMAGE_BASE_URL": "https://"}, {"IMAGE_BASE_URL": "https:// bad.example.test/v1"},
    {"IMAGE_MODEL": "bad\nmodel"}, {"IMAGE_MODEL": KEY},
    {"IMAGE_TIMEOUT": KEY}, {"IMAGE_TIMEOUT": "nan"}, {"IMAGE_TIMEOUT": "inf"},
    {"IMAGE_TIMEOUT": "0"}, {"IMAGE_TIMEOUT": "-1"}, {"IMAGE_TIMEOUT": "60.1"},
    {"IMAGE_RESPONSE_FORMAT": "url"},
])
def test_invalid_config_is_sanitized(changes):
    with pytest.raises(ImageGenError) as caught:
        ImageGenConfig.from_env(environment(**changes))
    assert str(caught.value) == "image_configuration_invalid"
    assert KEY not in str(caught.value)


@pytest.mark.parametrize("changes", [
    {"IMAGE_API_KEY": KEY + "-changed"}, {"IMAGE_MODEL": "changed-model"},
    {"IMAGE_BASE_URL": "https://changed.example.test/v1"},
    {"IMAGE_RESPONSE_FORMAT": "b64_json"}, {"IMAGE_TIMEOUT": "30"},
])
def test_private_binding_changes(changes):
    original = ImageGenConfig.from_env(environment())
    changed = ImageGenConfig.from_env(environment(**changes))
    assert original._binding_fingerprint() != changed._binding_fingerprint()
    assert len(original._binding_fingerprint()) == 64
    assert "_binding_fingerprint" not in original.public_config()


def test_direct_config_strict_timeout_and_host_only_image():
    for value in (True, "60", float("nan"), 61):
        with pytest.raises(ImageGenError):
            ImageGenConfig("openai_compatible", KEY, "https://images.example.test", "mock", timeout=value)
    metadata = {"sha256": "synthetic"}
    image = GeneratedImage(b"DO-NOT-LOG-RAW", metadata)
    metadata["sha256"] = "changed"
    assert "DO-NOT-LOG-RAW" not in repr(image)
    assert image.metadata["sha256"] == "synthetic"
    with pytest.raises(TypeError):
        image.metadata["sha256"] = "forbidden"
