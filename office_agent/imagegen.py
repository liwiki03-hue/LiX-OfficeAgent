"""P3 configuration foundation only; generation and save are not wired yet.

Configuration is lazy and separate from chat configuration. Raw images are host-only;
never return GeneratedImage through a model tool or serialize it into evidence.
"""
from __future__ import annotations

import hashlib
import json
import math
import os
import re
from dataclasses import dataclass, field
from types import MappingProxyType
from typing import Literal, Mapping, Protocol
from urllib.parse import urlsplit

from pydantic import BaseModel, ConfigDict, Field, StrictInt, StrictStr, field_validator


class ImageGenError(ValueError):
    """Stable public error code: never include configuration/provider exceptions."""

    def __init__(self, code: str = "image_configuration_invalid", *,
                 stage: str | None = None, detail_code: str | None = None,
                 request_state: str | None = None, provider_code: str | None = None) -> None:
        self.code = code
        self.stage = stage
        self.detail_code = detail_code
        self.request_state = request_state
        self.provider_code = provider_code
        super().__init__(code)


class ImageRequest(BaseModel):
    model_config = ConfigDict(extra="forbid", frozen=True)
    prompt: StrictStr = Field(min_length=1, max_length=4000)
    size: Literal["1024x1024", "1024x1536", "1536x1024"] = "1024x1024"
    style: Literal["natural", "vivid"] = "natural"
    n: StrictInt = Field(default=1, ge=1, le=1)

    @field_validator("prompt")
    @classmethod
    def validate_prompt(cls, value: str) -> str:
        if not value.strip() or any(0xD800 <= ord(char) <= 0xDFFF for char in value):
            raise ValueError("invalid_image_prompt")
        return value

    @property
    def effective_prompt(self) -> str:
        # Generic compatible APIs do not consistently accept the DALL-E style field.
        # The exact prefix MUST be displayed before any eventual cost approval.
        return f"Requested visual style: {self.style}.\n{self.prompt}"


def _plain(value: object, limit: int) -> bool:
    return (type(value) is str and 0 < len(value) <= limit
            and value == value.strip()
            and not any(ord(c) < 32 or ord(c) == 127 or 0xD800 <= ord(c) <= 0xDFFF for c in value))


@dataclass(frozen=True)
class ImageGenConfig:
    provider: str
    api_key: str = field(repr=False)
    base_url: str
    model: str
    response_format: Literal["auto", "b64_json"] = "auto"
    timeout: float = 60.0

    def __post_init__(self) -> None:
        try:
            valid = (self.provider == "openai_compatible" and _plain(self.api_key, 4096)
                     and _plain(self.base_url, 2048) and _plain(self.model, 128)
                     and re.fullmatch(r"[A-Za-z0-9][A-Za-z0-9._:/-]{0,127}", self.model)
                     and self.response_format in {"auto", "b64_json"}
                     and type(self.timeout) in {float, int}
                     and math.isfinite(self.timeout) and 0 < self.timeout <= 60)
            parts = urlsplit(self.base_url)
            valid = (valid and parts.scheme == "https" and bool(parts.hostname)
                     and parts.username is None and parts.password is None
                     and not parts.query and not parts.fragment
                     and "?" not in self.base_url and "#" not in self.base_url
                     and "\\" not in self.base_url and not any(c.isspace() for c in self.base_url)
                     and (parts.port is None or 1 <= parts.port <= 65535))
            if self.api_key in (self.provider + self.base_url + self.model + self.response_format):
                valid = False
        except Exception:
            valid = False
        if not valid:
            raise ImageGenError()

    @classmethod
    def from_env(cls, env: Mapping[str, str] | None = None) -> "ImageGenConfig":
        if env is None:
            from dotenv import load_dotenv
            load_dotenv(override=False)
            env = os.environ
        required = ("IMAGE_PROVIDER", "IMAGE_API_KEY", "IMAGE_BASE_URL", "IMAGE_MODEL")
        try:
            values = [env.get(name, "") for name in required]
            if not all(type(value) is str and bool(value) for value in values):
                raise ImageGenError("image_configuration_missing")
            return cls(*values, response_format=env.get("IMAGE_RESPONSE_FORMAT", "auto"),
                       timeout=float(env.get("IMAGE_TIMEOUT", "60")))
        except ImageGenError:
            raise
        except Exception:
            raise ImageGenError() from None

    def public_config(self) -> dict:
        return {"provider": self.provider, "base_url": self.base_url, "model": self.model,
                "response_format": self.response_format, "timeout": self.timeout}

    def _binding_fingerprint(self) -> str:
        # Host-only binding includes credentials; never emit this value in public metadata.
        body = {**self.public_config(), "api_key": self.api_key}
        return hashlib.sha256(json.dumps(body, sort_keys=True, separators=(",", ":")).encode()).hexdigest()


@dataclass(frozen=True)
class GeneratedImage:
    """Host-only raster. repr=False is not permission to put this in JSON/evidence."""
    raw: bytes = field(repr=False)
    metadata: Mapping = field(default_factory=dict)

    def __post_init__(self) -> None:
        object.__setattr__(self, "metadata", MappingProxyType(dict(self.metadata)))


class ImageGenProvider(Protocol):
    def generate(self, request: ImageRequest, config: ImageGenConfig) -> bytes:
        """Host invocation only, after an independent cloud/cost approval."""
        ...
