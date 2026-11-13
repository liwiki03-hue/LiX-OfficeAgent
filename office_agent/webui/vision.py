"""Explicit, bounded image-understanding requests; no agent tools or document writes."""
from __future__ import annotations

import base64
import io
import stat
from pathlib import Path

from PIL import Image, UnidentifiedImageError

from ..documents import (extract_document_image, read_excel_structure,
                         read_word_structure)
from ..writes import digest

MAX_BYTES = 10_000_000
MAX_PIXELS = 40_000_000


class VisionError(ValueError):
    pass


def _check_local(path: Path) -> bytes:
    for candidate in (path, *path.parents):
        info = candidate.lstat()
        if candidate.is_symlink() or getattr(info, "st_file_attributes", 0) & 0x400:
            raise VisionError("不接受链接或重解析点中的图片")
    before = path.stat()
    if not stat.S_ISREG(before.st_mode) or not 0 < before.st_size <= MAX_BYTES:
        raise VisionError("请选择不超过 10MB 的普通 PNG/JPEG 图片")
    with path.open("rb") as stream:
        raw = stream.read(MAX_BYTES + 1)
    after = path.stat()
    if (len(raw) != before.st_size or before.st_size != after.st_size or
            before.st_mtime_ns != after.st_mtime_ns or before.st_ino != after.st_ino):
        raise VisionError("图片在读取期间发生变化")
    return raw


def clean_image(raw: bytes) -> str:
    """Decode and re-encode one still image, dropping EXIF and embedded metadata."""
    if not isinstance(raw, bytes) or not 0 < len(raw) <= MAX_BYTES:
        raise VisionError("图片超过 10MB 限制")
    try:
        with Image.open(io.BytesIO(raw)) as image:
            if image.format not in {"PNG", "JPEG"} or getattr(image, "n_frames", 1) != 1:
                raise VisionError("仅支持单帧 PNG/JPEG")
            width, height = image.size
            if not 0 < width <= 32768 or not 0 < height <= 32768 or width * height > MAX_PIXELS:
                raise VisionError("图片像素超限")
            image.load()
            pixels = image.convert("RGB")
            clean = Image.new("RGB", image.size)
            clean.paste(pixels)
            buffer = io.BytesIO()
            clean.save(buffer, format="JPEG", quality=80, optimize=False)
            encoded = buffer.getvalue()
            if len(encoded) > MAX_BYTES:
                raise VisionError("图片重新编码后超过 10MB")
    except (OSError, ValueError, UnidentifiedImageError, Image.DecompressionBombError) as exc:
        if isinstance(exc, VisionError):
            raise
        raise VisionError("图片无法安全解码") from None
    return "data:image/jpeg;base64," + base64.b64encode(encoded).decode("ascii")


def document_images(path: Path) -> list[dict[str, str]]:
    if path.suffix.lower() == ".docx":
        items = read_word_structure(path, limit=50)["images"]
        return [{"relationship": item["relationship"], "part": item["source_part"],
                 "label": f"正文图片 {index + 1} ({item['relationship']})"}
                for index, item in enumerate(items)]
    if path.suffix.lower() == ".xlsx":
        sheets = read_excel_structure(path, limit=50)["sheets"]
        items = [(sheet["title"], image) for sheet in sheets for image in sheet.get("images", [])]
        return [{"relationship": image["relationship"], "part": image["source_part"],
                 "label": f"{title} 图片 {index + 1} ({image['relationship']})"}
                for index, (title, image) in enumerate(items[:50])]
    raise VisionError("文档图片仅支持 DOCX/XLSX")


def document_data(path: Path, version: str, relationship: str, part: str) -> str:
    if digest(path) != version:
        raise VisionError("所选文档已改变；请重新选择图片")
    items = document_images(path)
    if not any(item["relationship"] == relationship and item["part"] == part for item in items):
        raise VisionError("图片不在当前所选文档的结构清单中")
    image = extract_document_image(path, relationship, part=part)
    if digest(path) != version:
        raise VisionError("提取图片时文档发生变化")
    return clean_image(image.raw)


def understand(model, data_url: str, question: str) -> str:
    from langchain_core.messages import HumanMessage
    message = HumanMessage(content=[{"type": "text", "text": question},
                                    {"type": "image_url", "image_url": {"url": data_url}}])
    reply = model.invoke([message])
    text = reply.content
    if not isinstance(text, str) or not text.strip():
        raise VisionError("视觉模型未返回可显示的文本")
    return text[:12000]
