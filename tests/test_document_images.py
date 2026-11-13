"""P2 real OOXML fixtures, bounded display channel; no user documents or Office/cloud."""
import base64
import hashlib
import io
import json
import zipfile
from types import SimpleNamespace
from xml.etree import ElementTree as ET

import pytest
from docx import Document
from openpyxl import Workbook
from openpyxl.drawing.image import Image as ExcelImage
from PIL import Image
from pydantic import ValidationError

from office_agent import documents
from office_agent.app import DocumentImageArgs, build_dispatcher, build_tools, select_tools, render_tool_result
from office_agent.webui import bridge
from office_agent.writes import Editor, digest


@pytest.fixture(params=["docx", "xlsx"])
def image_source(tmp_path, request):
    stream = io.BytesIO()
    Image.new("RGB", (3, 2), "blue").save(stream, format="PNG")
    raw = stream.getvalue()
    path = tmp_path / ("synthetic." + request.param)
    if request.param == "docx":
        doc = Document()
        doc.add_paragraph("Synthetic only, no user input")
        doc.add_picture(io.BytesIO(raw))
        doc.save(path)
        location = documents.read_word_structure(path)["images"][0]
    else:
        book = Workbook()
        book.active.add_image(ExcelImage(io.BytesIO(raw)), "B3")
        book.save(path)
        location = documents._xlsx_sheet_images(path)["Sheet"][0]
    return path, raw, {"part": location["source_part"], "relationship": location["relationship"]}


def rewrite(path, change):
    with zipfile.ZipFile(path) as archive:
        parts = {entry.filename: archive.read(entry) for entry in archive.infolist()}
    change(parts)
    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        for name, blob in parts.items():
            archive.writestr(name, blob)


def alter_relation(path, args, **attributes):
    directory, _, name = args["part"].rpartition("/")
    rels = directory + "/_rels/" + name + ".rels"
    def change(parts):
        root = ET.fromstring(parts[rels])
        relation = next(node for node in root if node.get("Id") == args["relationship"])
        for key, value in attributes.items():
            relation.set(key, value)
        parts[rels] = ET.tostring(root)
    rewrite(path, change)


def test_locate_extract_display_preserves_source_and_model_metadata(image_source, tmp_path):
    path, raw, args = image_source
    before = digest(path)
    image = documents.extract_document_image(path, **args)
    assert image.raw == raw
    assert image.metadata["bytes_decoded"] is True
    assert image.metadata["width"] == 3 and image.metadata["height"] == 2
    assert image.metadata["format"] == "PNG"
    assert image.metadata["sha256"] == hashlib.sha256(raw).hexdigest()
    frames = []
    dispatch = build_dispatcher(Editor(tmp_path), path,
                                on_image=lambda image: frames.extend(bridge._document_image_display(image)))
    result = dispatch("read_document_image", {"path": str(path), **args})
    assert result["preview_available"] is True
    assert len(frames) == 1
    assert base64.b64decode(frames[0]["data_url"].split(",")[1]) == raw
    assert "data:" not in json.dumps(result) and "raw" not in result
    assert "元数据已解码" in render_tool_result("read_document_image", result)
    assert digest(path) == before
    assert not (tmp_path / ".office_agent" / "imagegen").exists()
    app = bridge.WebviewApp()
    app.path = path
    plan = SimpleNamespace(access="read", tool="read_document_image",
                           arguments={"path": str(path), **args}, file_versions={str(path): before})
    assert app._read_scope(plan) == path.resolve()
    app._read_authorized_path = path.resolve()
    assert app._confirm_read(plan) is True
    plan.access = "write"
    assert app._read_scope(plan) is None and app._confirm_read(plan) is False
    assert digest(path) == before


def test_no_pillow_is_honest_and_still_display_only(image_source, monkeypatch):
    path, raw, args = image_source
    monkeypatch.setattr(documents, "_PillowImage", None)
    image = documents.extract_document_image(path, **args)
    assert image.raw == raw and image.metadata["bytes_decoded"] is False
    assert "width" not in image.metadata and "height" not in image.metadata and "format" not in image.metadata
    assert len(bridge._document_image_display(image)) == 1
    assert "未安装 Pillow" in image.metadata["note"]


@pytest.mark.parametrize("attributes", [
    {"TargetMode": "External", "Target": "https://example.invalid/private.png"},
    {"TargetMode": "external"},
    {"Target": "../../../../outside.png"},
    {"Target": "file:///private.png"},
    {"Target": "//server/private.png"},
    {"Target": "../media/%2e%2e/private.png"},
    {"Type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"},
])
def test_relationship_attacks_refused(image_source, attributes):
    path, _, args = image_source
    alter_relation(path, args, **attributes)
    before = digest(path)
    with pytest.raises(documents.UnsupportedDocumentError):
        documents.extract_document_image(path, **args)
    assert digest(path) == before


def test_unembedded_relationship_and_wrong_part_refused(image_source):
    path, _, args = image_source
    with pytest.raises(documents.UnsupportedDocumentError):
        documents.extract_document_image(path, "rId999", part=args["part"])
    with pytest.raises(documents.UnsupportedDocumentError):
        documents.extract_document_image(path, args["relationship"], part="../../private.xml")


def test_oversized_image_rejected_before_opening_media(image_source, monkeypatch):
    path, _, args = image_source
    target = documents.extract_document_image(path, **args).metadata["target_part"]
    rewrite(path, lambda parts: parts.update({target: b"x" * 10_000_001}))
    opened = []
    original = zipfile.ZipFile.open
    def tracked(self, name, *a, **kw):
        opened.append(name.filename if isinstance(name, zipfile.ZipInfo) else name)
        return original(self, name, *a, **kw)
    monkeypatch.setattr(zipfile.ZipFile, "open", tracked)
    with pytest.raises(documents.UnsupportedDocumentError):
        documents.extract_document_image(path, **args)
    assert target not in opened


def test_duplicate_package_part_refused(image_source):
    path, _, args = image_source
    with zipfile.ZipFile(path, "a") as archive:
        with pytest.warns(UserWarning, match="Duplicate name"):
            archive.writestr(args["part"], b"ambiguous")
    with pytest.raises(documents.UnsupportedDocumentError):
        documents.extract_document_image(path, **args)


def test_malformed_raster_refused(image_source):
    path, _, args = image_source
    target = documents.extract_document_image(path, **args).metadata["target_part"]
    rewrite(path, lambda parts: parts.update({target: b"\x89PNG\r\n\x1a\ncorrupt"}))
    with pytest.raises(documents.UnsupportedDocumentError):
        documents.extract_document_image(path, **args)


def test_oversized_directory_refused_before_zipfile_allocation(tmp_path, monkeypatch):
    path = tmp_path / "synthetic.docx"
    with zipfile.ZipFile(path, "w") as archive:
        for i in range(2001):
            archive.writestr(f"entry{i}", b"")
    def forbidden(*args, **kwargs):
        raise AssertionError("ZipFile must not allocate oversized directory")
    monkeypatch.setattr(zipfile, "ZipFile", forbidden)
    with pytest.raises(documents.UnsupportedDocumentError, match="directory"):
        documents.extract_document_image(path, "rId1")


def test_xml_source_and_relationship_roots_refused(image_source):
    path, _, args = image_source
    directory, _, name = args["part"].rpartition("/")
    rels = directory + "/_rels/" + name + ".rels"
    for part in (args["part"], rels):
        original = path.read_bytes()
        def change(parts):
            root = ET.fromstring(parts[part])
            root.tag = "invalidRoot"
            parts[part] = ET.tostring(root)
        rewrite(path, change)
        with pytest.raises(documents.UnsupportedDocumentError):
            documents.extract_document_image(path, **args)
        path.write_bytes(original)


def test_dispatch_rejects_oversized_image_source_before_hash(image_source, tmp_path, monkeypatch):
    from office_agent import app
    path, _, args = image_source
    dispatch = build_dispatcher(Editor(tmp_path), path, expected_version="prebound")
    monkeypatch.setattr(documents, "_MAX_IMAGE_PACKAGE_BYTES", 1)
    def forbidden(*args):
        raise AssertionError("Image request must be rejected before hashing oversized source")
    monkeypatch.setattr(app, "digest", forbidden)
    with pytest.raises(documents.UnsupportedDocumentError):
        dispatch("read_document_image", {"path": str(path), **args})


def test_schema_tiers_and_stale_version_guard(image_source, tmp_path):
    path, _, args = image_source
    with pytest.raises(ValidationError):
        DocumentImageArgs(path=str(path), **args, confirmed=True)
    with pytest.raises(ValidationError):
        DocumentImageArgs(path=str(path), relationship=args["relationship"], part="/private.xml")
    tools = build_tools((path,))
    tool = next(t for t in tools if t.name == "read_document_image")
    assert tool.access == "read" and tool.allowed_paths == (path,)
    for kind in ("word", "spreadsheet"):
        assert tool in select_tools(tools, kind)
    assert tool not in select_tools(tools, "legacy")
    dispatch = build_dispatcher(Editor(tmp_path), path)
    path.write_bytes(path.read_bytes() + b"changed")
    with pytest.raises(ValueError, match="changed"):
        dispatch("read_document_image", {"path": str(path), **args})
