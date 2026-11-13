"""Synthetic image-understanding tests; never use a real account or provider."""
import base64
import io
from pathlib import Path
from types import SimpleNamespace

import pytest
from PIL import Image

from office_agent.documents import DocumentImage
from office_agent.webui import vision
from office_agent.webui.bridge import Api, WebviewApp
from office_agent.writes import digest


def png_bytes():
    buffer = io.BytesIO()
    image = Image.new("RGB", (2, 3), "#d42b2b")
    image.save(buffer, format="PNG", exif=b"private-metadata")
    return buffer.getvalue()


class Auth:
    def __init__(self, active=True):
        self.active = active

    def status(self):
        return {"authenticated": self.active}


class Model:
    def __init__(self):
        self.calls = []

    def invoke(self, messages):
        self.calls.append(messages)
        return SimpleNamespace(content="图中是一个红色方块")


def test_clean_image_removes_metadata_and_limits_formats():
    data = vision.clean_image(png_bytes())
    assert data.startswith("data:image/jpeg;base64,")
    encoded = base64.b64decode(data.partition(",")[2])
    with Image.open(io.BytesIO(encoded)) as image:
        assert image.size == (2, 3)
        assert not image.getexif()
    with pytest.raises(vision.VisionError):
        vision.clean_image(b"<svg/>")
    with pytest.raises(vision.VisionError):
        vision.clean_image(b"\x00" * (vision.MAX_BYTES + 1))


def test_local_file_rejects_oversize_and_links(tmp_path, monkeypatch):
    path = tmp_path / "picture.png"
    path.write_bytes(png_bytes())
    assert vision._check_local(path) == path.read_bytes()
    path.write_bytes(b"x" * (vision.MAX_BYTES + 1))
    with pytest.raises(vision.VisionError):
        vision._check_local(path)
    link = tmp_path / "linked.png"
    link.write_bytes(png_bytes())
    monkeypatch.setattr(Path, "is_symlink", lambda self: self == link)
    with pytest.raises(vision.VisionError):
        vision._check_local(link)


def test_document_version_and_membership_checked_before_extraction(tmp_path, monkeypatch):
    path = tmp_path / "selected.docx"
    path.write_bytes(b"synthetic anchor")
    monkeypatch.setattr(vision, "document_images", lambda _: [
        {"relationship": "rId1", "part": "word/document.xml", "label": "one"}])
    calls = []
    def extract(*args, **kwargs):
        calls.append((args, kwargs))
        return DocumentImage(png_bytes(), {"mime_type": "image/png"})
    monkeypatch.setattr(vision, "extract_document_image", extract)
    version = digest(path)
    assert vision.document_data(path, version, "rId1", "word/document.xml").startswith("data:image/jpeg;")
    with pytest.raises(vision.VisionError):
        vision.document_data(path, version, "rId2", "word/document.xml")
    path.write_bytes(b"changed source")
    with pytest.raises(vision.VisionError):
        vision.document_data(path, version, "rId1", "word/document.xml")
    assert len(calls) == 1


def test_model_receives_one_image_and_question_only():
    model = Model()
    data = vision.clean_image(png_bytes())
    assert "红色方块" in vision.understand(model, data, "描述图片")
    payload = model.calls[0][0].content
    assert payload[0] == {"type": "text", "text": "描述图片"}
    assert payload[1] == {"type": "image_url", "image_url": {"url": data}}


def test_api_denies_without_auth_or_model_configuration(tmp_path, monkeypatch):
    app = WebviewApp()
    path = tmp_path / "selected.docx"
    path.write_bytes(b"fixture")
    app.path = path
    app.editor = SimpleNamespace(space=SimpleNamespace(file=lambda p: Path(p)))
    auth = Auth(False)
    api = Api(app, auth=auth, settings_root=tmp_path)
    assert api.list_vision_images()["reason"] == "not_authenticated"
    assert api.understand_image("document", "描述", "rId1", "word/document.xml", digest(path))["reason"] == "not_authenticated"
    auth.active = True
    monkeypatch.setattr(vision, "document_images", lambda _: [
        {"relationship": "rId1", "part": "word/document.xml", "label": "one"}])
    monkeypatch.setattr(vision, "extract_document_image", lambda *args, **kwargs:
                        DocumentImage(png_bytes(), {"mime_type": "image/png"}))
    inventory = api.list_vision_images()
    assert inventory["ok"] and inventory["version"] == digest(path)
    monkeypatch.setattr("dotenv.load_dotenv", lambda *args, **kwargs: None)
    monkeypatch.delenv("OFFICE_VISION_MODEL", raising=False)
    result = api.understand_image("document", "描述", "rId1", "word/document.xml", digest(path))
    assert result["reason"] == "vision_invalid" and "OFFICE_VISION_MODEL" in result["error"]
    assert api.understand_image("document", "描述", "rId1", "word/document.xml", "old")["reason"] == "vision_invalid"
    assert app.poll() == [] and app._history_events == [] and path.read_bytes() == b"fixture"


def test_api_configured_fake_model_read_only(tmp_path, monkeypatch):
    app = WebviewApp()
    path = tmp_path / "selected.docx"
    path.write_bytes(b"fixture")
    app.path = path
    app.editor = SimpleNamespace(space=SimpleNamespace(file=lambda p: Path(p)))
    api = Api(app, auth=Auth(), settings_root=tmp_path)
    monkeypatch.setattr(vision, "document_images", lambda _: [
        {"relationship": "rId1", "part": "word/document.xml", "label": "one"}])
    monkeypatch.setattr(vision, "extract_document_image", lambda *args, **kwargs:
                        DocumentImage(png_bytes(), {"mime_type": "image/png"}))
    monkeypatch.setattr("dotenv.load_dotenv", lambda *args, **kwargs: None)
    monkeypatch.setenv("OFFICE_VISION_MODEL", "deepseek-flash")
    monkeypatch.setattr("main.read_api_key", lambda provider: "synthetic-test-key")
    model = Model()
    config = {}
    def fake_client(**kwargs):
        config.update(kwargs)
        return model
    monkeypatch.setattr("langchain_openai.ChatOpenAI", fake_client)
    result = api.understand_image("document", "解释图像", "rId1", "word/document.xml", digest(path))
    assert result == {"ok": True, "answer": "图中是一个红色方块"}
    assert config["model"] == "deepseek-flash" and config["max_retries"] == 0
    assert config["api_key"].get_secret_value() == "synthetic-test-key"
    assert len(model.calls) == 1
    assert app.poll() == [] and app._history_events == [] and path.read_bytes() == b"fixture"
