"""Safe document classification and read-only inspection helpers.

The module deliberately does not mutate source files or execute embedded content.
Legacy OLE formats are classified from metadata and reported as unsupported unless
an explicit COM reader is supplied by the caller (no COM automation is bundled).
"""
from __future__ import annotations

import hashlib
import io
import stat
import struct
import warnings
import json
import mimetypes
import os
import re
import zipfile
from concurrent.futures import ThreadPoolExecutor
from dataclasses import dataclass, field, asdict
from pathlib import Path
from typing import Any, Callable, Iterable, Mapping, Protocol, Sequence
from xml.etree import ElementTree as ET

try:
    from docx import Document as _DocxDocument
    from docx.table import Table as _DocxTable
except ImportError:  # optional until a DOCX is inspected
    _DocxDocument = None
    _DocxTable = None
try:
    import openpyxl
    from openpyxl.utils import column_index_from_string, get_column_letter
except ImportError:  # optional until an XLSX is inspected
    openpyxl = None
    get_column_letter = None
    column_index_from_string = None


OLE_MAGIC = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"
OOXML_EXTENSIONS = {".docx", ".xlsx", ".xlsm", ".xltx", ".xltm", ".pptx"}
LEGACY_EXTENSIONS = {".doc", ".xls", ".ppt"}
_MAX_STRUCTURE_TEXT = 1024  # Per-item cap for structure reads; body text has its own tools.
_MAX_OUTLINE_HEADINGS = 100  # Heading outline cap; a longer outline is flagged, not dumped.
_MAX_TABLE_SUMMARIES = 50    # Table shape summaries; content is fetched separately.
_MIN_TEXT_BLOCKS = 5         # Floor for read_word_text: fewer blocks cannot describe a document.
_MIN_TEXT_BUDGET = 1000      # Floor for read_word_text text budget, in characters.
_MAX_STRUCTURE_LIMIT = 100   # Hard cap for one bounded structure read, whatever the caller asks.
_MAX_STRUCTURE_SHEETS = 20   # Workbooks are detailed sheet by sheet; more sheets are flagged.
_MAX_REVISION_PARTS = 20     # DOCX parts scanned for tracked changes and content controls.
_MAX_TABLE_DEPTH = 3         # Nested tables are expanded to this depth, then flagged.
_MAX_TABLE_ISSUES = 20       # Consistency issues listed; the total count is always reported.
_BATCH_MAX_FILES = 20        # Hard cap on one batch run; a larger directory is flagged.
_BATCH_MAX_WORKERS = 4       # Independent read-only siblings; never parallelize path admission.
_BATCH_MAX_PER_FILE = 20     # Hard cap on items reported per file.
_BATCH_SUFFIXES = (".docx", ".xlsx")        # Only what the Python readers handle safely.
_BATCH_LEGACY_SUFFIXES = (".doc", ".xls")   # Recognised, deliberately never parsed here.
# Our own published copies and Office lock files are not batch inputs.
_BATCH_OUTPUT = re.compile(r"\.agent-[0-9a-f]{8}\.", re.IGNORECASE)
_BATCH_TEXT = 120            # Characters kept per batch item (a result panel, not the file).
_MAX_NAMED_REFERENCE = 200   # Named-range targets and chart references are clipped, then flagged.
_MAX_STRUCTURE_PART_BYTES = 20_000_000  # Parts above this are listed as unscanned, never as empty.
_WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
# Read-only inventory includes revisions deliberately refused by the write channel.
_WORD_REVISION_DETAIL_MARKS = frozenset({
    "ins", "del", "rPrChange", "pPrChange", "moveFrom", "moveTo",
    "moveFromRangeStart", "moveFromRangeEnd", "moveToRangeStart", "moveToRangeEnd",
    "cellIns", "cellDel", "cellMerge", "tblPrChange", "tblGridChange", "tcPrChange",
    "trPrChange", "sectPrChange", "numberingChange", "customXmlInsRangeStart",
    "customXmlInsRangeEnd", "customXmlDelRangeStart", "customXmlDelRangeEnd",
    "customXmlMoveFromRangeStart", "customXmlMoveFromRangeEnd",
    "customXmlMoveToRangeStart", "customXmlMoveToRangeEnd",
})
# Content-control type markers inside w:sdtPr, in the order Word writes them.
_SDT_KINDS = ("text", "richText", "date", "dropDownList", "comboBox", "checkbox", "picture",
              "buildingBlockGallery", "docPartObj", "group", "citation", "bibliography")


@dataclass(frozen=True)
class DocumentClassification:
    path: str
    extension: str
    kind: str
    format: str
    readable: bool
    protected: bool = False
    macro_enabled: bool = False
    external_relationships: bool = False
    suspicious_features: tuple[str, ...] = ()
    metadata: Mapping[str, Any] = field(default_factory=dict)
    limitations: tuple[str, ...] = ()

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


@dataclass(frozen=True)
class OCRResult:
    text: str
    source: str
    confidence: float | None = None
    language: str | None = None
    metadata: Mapping[str, Any] = field(default_factory=dict)


class OCRProvider(Protocol):
    """Pluggable OCR contract. Providers must not mutate the input document."""
    def recognize(self, image: Any, *, source: str, **kwargs: Any) -> OCRResult: ...


class UnsupportedDocumentError(ValueError):
    pass


def _has_dde_field(xml: bytes) -> bool:
    """Only treat Word field instructions as DDE, not incidental template bytes."""
    root = ET.fromstring(xml)
    instructions: list[str] = []
    for node in root.iter():
        local_name = node.tag.rsplit("}", 1)[-1].lower()
        if local_name == "instrtext":
            instructions.append("".join(node.itertext()))
        elif local_name == "fldsimple":
            instructions.extend(value for key, value in node.attrib.items()
                                if key.rsplit("}", 1)[-1].lower() == "instr")
    return bool(re.search(r"\bDDE(?:AUTO)?\b", " ".join(instructions), re.I))


def _has_active_protection(xml: bytes) -> bool:
    """Ignore inert OOXML protection elements; inspect their activation flags."""
    for node in ET.fromstring(xml).iter():
        name = node.tag.rsplit("}", 1)[-1].lower()
        if name not in {"documentprotection", "workbookprotection", "sheetprotection"}:
            continue
        attrs = {key.rsplit("}", 1)[-1].lower(): value.strip().lower()
                 for key, value in node.attrib.items()}
        def enabled(key: str) -> bool:
            return attrs.get(key, "") not in {"", "0", "false", "off"}
        if name == "documentprotection" and (enabled("enforcement") or
                any(enabled(key) for key in ("hash", "hashvalue", "password"))):
            return True
        if name == "workbookprotection" and (any(enabled(key) for key in
                ("lockstructure", "lockwindows", "lockrevision", "workbookpassword",
                 "revisionspassword", "workbookhashvalue", "revisionshashvalue"))):
            return True
        if name == "sheetprotection" and ("sheet" not in attrs or enabled("sheet") or
                any(enabled(key) for key in ("password", "hashvalue"))):
            return True
    return False


def _zip_report(path: Path) -> tuple[bool, bool, bool, list[str], dict[str, Any]]:
    """Inspect package names and relationship XML without opening application parts."""
    suspicious: list[str] = []
    metadata: dict[str, Any] = {}
    if not zipfile.is_zipfile(path):
        return False, False, False, ["invalid_zip"], metadata
    try:
        with zipfile.ZipFile(path) as zf:
            names = zf.namelist()
            required = {".docx": "word/document.xml", ".docm": "word/document.xml",
                        ".xlsx": "xl/workbook.xml", ".xlsm": "xl/workbook.xml",
                        ".xltx": "xl/workbook.xml", ".xltm": "xl/workbook.xml", ".pptx": "ppt/presentation.xml"}
            if path.suffix.lower() in required and (
                    "[Content_Types].xml" not in names or required[path.suffix.lower()] not in names):
                suspicious.append("missing_ooxml_core_part")
            metadata["part_count"] = len(names)
            metadata["uncompressed_bytes"] = sum(i.file_size for i in zf.infolist())
            if metadata["uncompressed_bytes"] > 512 * 1024 * 1024 or len(names) > 10000:
                return False, False, False, ["possible_zip_bomb", "inspection_incomplete"], metadata
            if metadata["uncompressed_bytes"] / max(path.stat().st_size, 1) > 1000:
                return False, False, False, ["possible_zip_bomb", "inspection_incomplete"], metadata
            protected = any(n.lower() in {"encryptedpackage", "encryptedpackageinfo"} for n in names)
            macro = any(n.lower().endswith("vbaproject.bin") for n in names)
            rel_parts = [n for n in names if n.lower().endswith(".rels")]
            external = False
            for name in rel_parts:
                if zf.getinfo(name).file_size > 2_000_000:
                    suspicious.append("large_relationships_uninspected")
                    continue
                try:
                    data = zf.read(name)
                except (KeyError, RuntimeError, zipfile.BadZipFile):
                    continue
                text = data.decode("utf-8", "ignore")
                if "TargetMode=\"External\"" in text or "TargetMode='External'" in text:
                    external = True
                    suspicious.append("external_relationship")
            lowered = {n.lower() for n in names}
            metadata["images"] = any("/media/" in n for n in lowered)
            metadata["charts"] = any("/charts/" in n for n in lowered)
            metadata["tables"] = False
            metadata["formulas"] = False
            for part in names:
                low = part.lower()
                if low == "word/document.xml" or re.fullmatch(r"xl/worksheets/sheet\d+\.xml", low):
                    # Bound scanning of big parts and report uncertainty rather than claiming absent.
                    info = zf.getinfo(part)
                    if info.file_size > 20_000_000:
                        suspicious.append("large_xml_part_not_fully_inspected")
                        continue
                    blob = zf.read(part).lower()
                    metadata["tables"] |= b"<w:tbl" in blob or b"<tableparts" in blob
                    metadata["formulas"] |= b"<m:omath" in blob or bool(re.search(rb"<f(?:\s|>)", blob))
                    if low == "word/document.xml":
                        for token in (b"<w:sdt", b"<w:ins", b"<w:del", b"<w:fldchar", b"<w:instrtext", b"<m:omath"):
                            if token in blob:
                                suspicious.append("word_advanced_xml:" + token.decode("ascii"))
                if low in {"word/settings.xml", "xl/workbook.xml"} or re.fullmatch(r"xl/worksheets/sheet\d+\.xml", low):
                    if zf.getinfo(part).file_size > 20_000_000:
                        suspicious.append("large_protection_part_uninspected")
                    else:
                        data = zf.read(part)
                        if b"protection" in data.lower():
                            try:
                                protected |= _has_active_protection(data)
                            except ET.ParseError:
                                suspicious.append("unparseable_protection_xml")
            for marker in ("/pivot", "/querytables/", "/connections", "/slicers/", "/threadedcomments/", "/customxml/", "/charts/", "/drawings/", "word/comments", "word/footnotes", "word/endnotes"):
                if any(marker in name for name in lowered):
                    suspicious.append("unsupported_part:" + marker)
            for marker, label in (("/activex/", "activex"), ("/embeddings/", "embedded_object"),
                                  ("/customui", "custom_ui"), ("externalLinks/".lower(), "external_link")):
                if any(marker in n for n in lowered):
                    suspicious.append(label)
            # DDE and script-like payloads are suspicious even if not active.
            for name in names:
                if name.lower().endswith((".xml", ".rels", ".bin")):
                    if zf.getinfo(name).file_size > 2_000_000:
                        suspicious.append("large_part_uninspected")
                        continue
                    try:
                        blob = zf.read(name).lower()
                    except Exception:
                        continue
                    if (b"dde" in blob and name.lower().startswith("word/")
                            and name.lower().endswith(".xml")):
                        try:
                            if _has_dde_field(blob):
                                suspicious.append("dde_marker")
                        except ET.ParseError:
                            suspicious.append("unparseable_word_xml")
                    if b"javascript:" in blob or b"wscript.shell" in blob:
                        suspicious.append("script_marker")
            ratio = metadata["uncompressed_bytes"] / max(path.stat().st_size, 1)
            metadata["compression_ratio"] = round(ratio, 2)
            if ratio > 1000 or metadata["uncompressed_bytes"] > 512 * 1024 * 1024:
                suspicious.append("possible_zip_bomb")
            return protected, macro, external, sorted(set(suspicious)), metadata
    except (OSError, zipfile.BadZipFile):
        return False, False, False, ["invalid_zip"], metadata


def classify_document(path: str | os.PathLike[str]) -> DocumentClassification:
    p = Path(path)
    ext = p.suffix.lower()
    metadata: dict[str, Any] = {"exists": p.exists(), "size": p.stat().st_size if p.exists() else None}
    if not p.exists() or not p.is_file():
        return DocumentClassification(str(p), ext, "unknown", "unknown", False, metadata=metadata,
                                      limitations=("file_not_found",))
    with p.open("rb") as stream:
        header = stream.read(8)
    if ext == '.pdf':
        return DocumentClassification(str(p), ext, 'pdf', 'pdf', header.startswith(b'%PDF-'),
                                      metadata=metadata, limitations=('pdf_parser_validation_on_read',))
    if ext in {".doc", ".xls"} or header == OLE_MAGIC:
        metadata["mime"] = mimetypes.guess_type(str(p))[0]
        return DocumentClassification(str(p), ext, "legacy", "ole_compound", False,
                                      metadata=metadata, limitations=("legacy_format_requires_com_or_specialized_reader",))
    if ext in OOXML_EXTENSIONS or zipfile.is_zipfile(p):
        protected, macro, external, suspicious, zmeta = _zip_report(p)
        metadata.update(zmeta)
        kind = "word" if ext in {".docx", ".docm"} else "spreadsheet" if ext in {".xlsx", ".xlsm", ".xltx", ".xltm"} else "presentation" if ext == '.pptx' else "ooxml"
        limitations = []
        if protected: limitations.append("encrypted_or_protected_package")
        if macro: limitations.append("macros_not_executed")
        if suspicious: limitations.append("suspicious_package_features_present")
        unsafe_to_open = {"invalid_zip", "missing_ooxml_core_part", "possible_zip_bomb", "inspection_incomplete"}
        return DocumentClassification(str(p), ext, kind, "ooxml", not protected and not unsafe_to_open.intersection(suspicious),
                                      protected, macro, external, tuple(suspicious), metadata, tuple(limitations))
    return DocumentClassification(str(p), ext, "unknown", "unknown", False, metadata=metadata,
                                  limitations=("unsupported_format",))


def _chunks(text: str, chunk_size: int, overlap: int) -> list[str]:
    if chunk_size <= 0 or overlap < 0 or overlap >= chunk_size:
        raise ValueError("chunk_size must be positive and overlap smaller than chunk_size")
    return [text[i:i + chunk_size] for i in range(0, max(len(text), 1), chunk_size - overlap)]


def _paragraph_info(paragraph: Any) -> dict[str, Any]:
    return {"text": paragraph.text, "style": paragraph.style.name if paragraph.style else None,
            "is_heading": bool(paragraph.style and paragraph.style.name.lower().startswith("heading"))}


# Image bytes are a display-only channel, never a model/tool/checkpoint payload.
try:
    from PIL import Image as _PillowImage
except ImportError:
    _PillowImage = None

_MAX_IMAGE_BYTES = 10_000_000
_MAX_IMAGE_PIXELS = 40_000_000
_MAX_IMAGE_PACKAGE_BYTES = 100_000_000


@dataclass(frozen=True)
class DocumentImage:
    raw: bytes = field(repr=False)
    metadata: Mapping[str, Any]


def _image_part_target(part: str, target: str) -> str:
    if not target or len(target) > 240 or any(c in target for c in "\\:%?#\x00") or target.startswith("//"):
        raise UnsupportedDocumentError("Unsafe image relationship target")
    pieces = [] if target.startswith("/") else part.split("/")[:-1]
    for piece in target.lstrip("/").split("/"):
        if piece == "..":
            if not pieces:
                raise UnsupportedDocumentError("Image target escapes package root")
            pieces.pop()
        elif piece not in {"", "."}:
            pieces.append(piece)
    name = "/".join(pieces)
    prefix = "word" if part.startswith("word/") else "xl"
    if not re.fullmatch(prefix + r"/media/[A-Za-z0-9_. -]{1,160}", name):
        raise UnsupportedDocumentError("Image target is outside package media")
    return name


def _image_xml(archive: zipfile.ZipFile, part: str) -> ET.Element:
    if archive.getinfo(part).file_size > 2_000_000:
        raise UnsupportedDocumentError("Image XML part exceeds 2MB bound")
    blob = archive.read(part)
    if b"\x00" in blob or b"<!DOCTYPE" in blob.upper() or b"<!ENTITY" in blob.upper():
        raise UnsupportedDocumentError("Image XML entities/unsupported encoding refused")
    root = ET.fromstring(blob)
    stack = [(root, 1)]
    count = 0
    while stack:
        node, depth = stack.pop()
        count += 1
        if depth > 256 or count > 100000:
            raise UnsupportedDocumentError("Image XML depth/node bound exceeded")
        stack.extend((child, depth + 1) for child in node)
    return root


def _image_source_preflight(path: Path):
    for candidate in (path, *path.parents):
        info = candidate.lstat()
        if candidate.is_symlink() or getattr(info, "st_file_attributes", 0) & 0x400:
            raise UnsupportedDocumentError("Linked image source refused")
    info = path.stat()
    if not stat.S_ISREG(info.st_mode) or info.st_size > _MAX_IMAGE_PACKAGE_BYTES:
        raise UnsupportedDocumentError("Image source exceeds 100MB package bound")
    return info


def _image_zip_preflight(path: Path, size: int) -> None:
    """Bound central-directory allocation BEFORE Python ZipFile parses it (no ZIP64)."""
    with path.open("rb") as stream:
        tail_size = min(size, 65557)
        stream.seek(size - tail_size)
        tail = stream.read(tail_size)
        position = tail.rfind(b"PK\x05\x06")
        if position < 0 or position + 22 > len(tail):
            raise UnsupportedDocumentError("Missing bounded ZIP directory")
        _, disk, directory_disk, on_disk, count, length, offset, comment = struct.unpack_from("<4s4H2IH", tail, position)
        if (disk or directory_disk or count != on_disk or count > 2000 or length > 2_000_000 or
                position + 22 + comment != len(tail) or offset + length != size - tail_size + position):
            raise UnsupportedDocumentError("Oversized/ZIP64/multidisk image directory refused")
        stream.seek(offset)
        directory = stream.read(length)
    cursor = seen = 0
    while cursor < len(directory):
        if seen >= 2000 or cursor + 46 > len(directory) or directory[cursor:cursor + 4] != b"PK\x01\x02":
            raise UnsupportedDocumentError("Malformed/oversized image directory")
        name, extra, comment = struct.unpack_from("<HHH", directory, cursor + 28)
        cursor += 46 + name + extra + comment
        seen += 1
    if cursor != len(directory) or seen != count:
        raise UnsupportedDocumentError("Ambiguous image directory")


def extract_document_image(path: str | os.PathLike[str], relationship: str,
                           *, part: str = "word/document.xml") -> DocumentImage:
    """Extract one bounded internal DOCX/XLSX image, without writing or external I/O.

    Callers still require Workspace authorization and the selected-file version. Raw
    bytes must go only to the display callback; return metadata to the model.
    """
    path = Path(path)
    if path.suffix.lower() not in {".docx", ".xlsx"}:
        raise UnsupportedDocumentError("DOCX/XLSX required for image extraction")
    before = _image_source_preflight(path)
    _image_zip_preflight(path, before.st_size)
    if not re.fullmatch(r"rId[0-9]{1,8}", relationship):
        raise UnsupportedDocumentError("Invalid image relationship")
    pattern = (r"word/(?:document|header[0-9]{1,8}|footer[0-9]{1,8})\.xml"
               if path.suffix.lower() == ".docx" else r"xl/drawings/drawing[0-9]{1,8}\.xml")
    if not re.fullmatch(pattern, part):
        raise UnsupportedDocumentError("Invalid image source part")
    try:
        with zipfile.ZipFile(path) as archive:
            entries = archive.infolist()
            names = [item.filename for item in entries]
            if (len(entries) > 2000 or len(names) != len(set(names)) or
                    sum(item.file_size for item in entries) > _MAX_IMAGE_PACKAGE_BYTES or
                    any(item.flag_bits & 1 for item in entries)):
                raise UnsupportedDocumentError("Ambiguous/encrypted/oversized image package refused")
            if any(item.file_size > 2_000_000 and item.filename.endswith((".xml", ".rels")) for item in entries):
                raise UnsupportedDocumentError("Image package XML exceeds extraction bounds")
            root = _image_xml(archive, part)
            for item in entries:
                if item.filename.endswith((".xml", ".rels")):
                    blob = archive.read(item)
                    if b"\x00" in blob or b"<!DOCTYPE" in blob.upper() or b"<!ENTITY" in blob.upper():
                        raise UnsupportedDocumentError("Unsafe image package XML refused before classification")
            classification = classify_document(path)
            if not classification.readable:
                raise UnsupportedDocumentError("Unreadable or risky image source refused")
            # Prove this relationship is actually embedded by an image drawing.
            word_ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            draw_ns = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
            picture_ns = "http://schemas.openxmlformats.org/drawingml/2006/picture"
            local_root = ("document" if part == "word/document.xml" else
                          "hdr" if "/header" in part else "ftr")
            expected_root = "{" + (word_ns if path.suffix.lower() == ".docx" else draw_ns) + "}" + (
                local_root if path.suffix.lower() == ".docx" else "wsDr")
            if root.tag != expected_root:
                raise UnsupportedDocumentError("Invalid image source XML root")
            parents = {child: parent for parent in root.iter() for child in parent}
            def drawing_image(node):
                if path.suffix.lower() == ".xlsx":
                    parent = parents.get(node)
                    picture = parents.get(parent)
                    anchor = parents.get(picture)
                    return (parent is not None and parent.tag == "{" + draw_ns + "}blipFill" and
                            picture is not None and picture.tag == "{" + draw_ns + "}pic" and
                            anchor is not None and anchor.tag in {"{" + draw_ns + "}" + name
                                for name in ("oneCellAnchor", "twoCellAnchor", "absoluteAnchor")})
                parent = parents.get(node)
                picture = parents.get(parent)
                if parent is None or parent.tag != "{" + picture_ns + "}blipFill" or picture is None or picture.tag != "{" + picture_ns + "}pic":
                    return False
                current = picture
                while current in parents:
                    current = parents[current]
                    if current.tag == "{" + word_ns + "}drawing":
                        return True
                return False
            embedded = {node.get("{" + _REL_NS + "}embed") for node in root.iter()
                        if node.tag == "{http://schemas.openxmlformats.org/drawingml/2006/main}blip" and drawing_image(node)}
            if relationship not in embedded:
                raise UnsupportedDocumentError("Relationship is not an embedded image in this part")
            directory, _, name = part.rpartition("/")
            rels = _image_xml(archive, f"{directory}/_rels/{name}.rels")
            if rels.tag != "{http://schemas.openxmlformats.org/package/2006/relationships}Relationships":
                raise UnsupportedDocumentError("Invalid image relationships XML root")
            matches = [node for node in rels if node.get("Id") == relationship]
            if len(matches) != 1:
                raise UnsupportedDocumentError("Missing/ambiguous image relationship")
            relation = matches[0]
            if (relation.tag != "{http://schemas.openxmlformats.org/package/2006/relationships}Relationship" or
                    relation.get("Type") != _REL_TYPE + "image" or
                    relation.get("TargetMode", "Internal") != "Internal"):
                raise UnsupportedDocumentError("External/non-image relationship refused")
            target = _image_part_target(part, relation.get("Target", ""))
            if archive.getinfo(target).file_size > _MAX_IMAGE_BYTES:
                raise UnsupportedDocumentError("Image exceeds 10MB bound")
            with archive.open(target) as stream:
                raw = stream.read(_MAX_IMAGE_BYTES + 1)
            if not raw or len(raw) > _MAX_IMAGE_BYTES:
                raise UnsupportedDocumentError("Empty/oversized image refused")
    except (zipfile.BadZipFile, ET.ParseError, KeyError, RuntimeError) as exc:
        raise UnsupportedDocumentError("Image package could not be read safely") from None
    after = path.stat()
    if (before.st_ino, before.st_size, before.st_mtime_ns) != (after.st_ino, after.st_size, after.st_mtime_ns):
        raise UnsupportedDocumentError("Image source changed during extraction")
    mime = ("image/png" if raw.startswith(b"\x89PNG\r\n\x1a\n") else
            "image/jpeg" if raw.startswith(b"\xff\xd8\xff") else "application/octet-stream")
    metadata = {"part": part, "relationship": relationship, "target_part": target,
                "byte_count": len(raw), "sha256": hashlib.sha256(raw).hexdigest(),
                "mime_type": mime, "bytes_decoded": False,
                "preview_supported": mime in {"image/png", "image/jpeg"},
                "note": "未安装 Pillow，未解码元数据；预览仍须通过前端尺寸与解码校验"}
    # Never allow APNG into the browser's unbounded animation decoder.
    if mime == "image/png":
        offset, chunks = 8, 0
        while offset + 12 <= len(raw):
            length = int.from_bytes(raw[offset:offset + 4], "big")
            kind = raw[offset + 4:offset + 8]
            if offset + 12 + length > len(raw) or chunks > 50000:
                raise UnsupportedDocumentError("Malformed/oversized PNG chunk structure")
            if kind == b"acTL":
                metadata.update(preview_supported=False, note="动画 PNG 不预览，仅返回有界来源元数据")
            offset += 12 + length
            chunks += 1
            if kind == b"IEND":
                break
    if _PillowImage is not None:
        try:
            with warnings.catch_warnings():
                warnings.simplefilter("error", _PillowImage.DecompressionBombWarning)
                with _PillowImage.open(io.BytesIO(raw)) as image:
                    width, height = image.size
                    if not (0 < width <= 32768 and 0 < height <= 32768 and width * height <= _MAX_IMAGE_PIXELS):
                        raise UnsupportedDocumentError("Image dimensions exceed decoding bounds")
                    if getattr(image, "n_frames", 1) != 1:
                        metadata["preview_supported"] = False
                    image_format = image.format
                    image.verify()
                metadata.update(bytes_decoded=True, width=width, height=height,
                                format=image_format, note="仅解码/验证图片元数据，不代表内容识别")
        except (OSError, ValueError, _PillowImage.DecompressionBombWarning,
                _PillowImage.DecompressionBombError) as exc:
            if mime in {"image/png", "image/jpeg"}:
                raise UnsupportedDocumentError("Image decoding/dimension validation failed") from None
            metadata.update(preview_supported=False, note="此图片格式无法安全解码或预览")
    if not metadata["preview_supported"]:
        metadata["note"] = "图片已提取；此格式/动画不支持当前前端安全预览"
    return DocumentImage(raw, metadata)


def _docx_image_drawings(doc: Any) -> list[dict[str, Any]]:
    """Return image drawings in document body paragraphs and table cells."""
    drawings: list[dict[str, Any]] = []
    ns = {
        "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
        "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    }

    def add_from_paragraph(paragraph: Any, location: dict[str, Any]) -> None:
        for drawing in paragraph._p.findall(".//w:drawing", ns):
            blip = drawing.find(".//a:blip", ns)
            rid = blip.get("{%s}embed" % ns["r"]) if blip is not None else None
            if not rid:
                continue
            placement = "inline" if drawing.find(".//wp:inline", ns) is not None else (
                "anchored" if drawing.find(".//wp:anchor", ns) is not None else None
            )
            drawings.append({"relationship": rid, "position": location, "placement": placement})

    for paragraph_index, paragraph in enumerate(doc.paragraphs):
        add_from_paragraph(paragraph, {"type": "paragraph", "index": paragraph_index})
    for table_index, table in enumerate(doc.tables):
        for row_index, row in enumerate(table.rows):
            for column_index, cell in enumerate(row.cells):
                location = {"type": "table_cell", "table": table_index,
                            "row": row_index, "column": column_index}
                for paragraph in cell.paragraphs:
                    add_from_paragraph(paragraph, location)
    return drawings


def ocr_docx_image(path: str | os.PathLike[str], relationship: str, ocr: OCRProvider) -> dict[str, Any]:
    """Opt-in OCR of one internal image; never writes recognized text into DOCX."""
    c = classify_document(path)
    if c.kind != "word" or not c.readable or Path(path).suffix.lower() != ".docx":
        raise UnsupportedDocumentError("only readable DOCX images can be recognized")
    if not re.fullmatch(r"rId[0-9]{1,8}", relationship):
        raise ValueError("relationship must be one bounded DOCX relationship ID")
    if _DocxDocument is None:
        raise UnsupportedDocumentError("python-docx is not installed")
    doc = _DocxDocument(path)
    rel = doc.part.rels.get(relationship)
    if rel is None or not rel.reltype.endswith("/image") or rel.is_external:
        raise ValueError("relationship does not identify an internal image")
    image = rel.target_part.blob
    if len(image) > 10_000_000:
        raise ValueError("image too large for opt-in OCR")
    source = f"{path}#relationship={relationship}"
    result = ocr.recognize(image, source=source)
    if not isinstance(result, OCRResult) or result.source != source or (
            result.confidence is not None and not 0 <= result.confidence <= 1):
        raise ValueError("OCR provider must return a matching source and confidence between 0 and 1")
    if len(result.text) > 4096 or len(str(result.metadata)) > 2048:
        raise ValueError("OCR result exceeds bounded tool output")
    return asdict(result)


def inspect_docx(path: str | os.PathLike[str], *, ocr: OCRProvider | None = None) -> dict[str, Any]:
    c = classify_document(path)
    if c.kind != "word" or c.protected or not c.readable:
        return {"classification": c.to_dict(), "readable": False, "limitations": list(c.limitations)}
    if _DocxDocument is None: raise UnsupportedDocumentError("python-docx is not installed")
    doc = _DocxDocument(path)
    paragraphs = [_paragraph_info(p) for p in doc.paragraphs]
    tables = [[[cell.text for cell in row.cells] for row in table.rows] for table in doc.tables]
    sections = []
    for section in doc.sections:
        sections.append({"header": [p.text for p in section.header.paragraphs],
                         "footer": [p.text for p in section.footer.paragraphs]})
    styles = sorted({p["style"] for p in paragraphs if p["style"]})
    comments: list[dict[str, Any]] = []
    try:
        with zipfile.ZipFile(path) as zf:
            if "word/comments.xml" in zf.namelist():
                root = ET.fromstring(zf.read("word/comments.xml"))
                ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
                for node in root.findall("w:comment", ns):
                    comments.append({"id": node.get("{%s}id" % ns["w"]),
                                     "author": node.get("{%s}author" % ns["w"]),
                                     "text": "".join(node.itertext())})
    except (ET.ParseError, zipfile.BadZipFile):
        comments = []
    drawing_info = _docx_image_drawings(doc)
    by_relationship: dict[str, list[dict[str, Any]]] = {}
    for item in drawing_info:
        by_relationship.setdefault(item["relationship"], []).append(item)
    images: list[dict[str, Any]] = []
    for rel in doc.part.rels.values():
        if "image" not in rel.reltype:
            continue
        image = {"relationship": rel.rId, "target": str(rel.target_ref),
                 "source": "document", "source_part": str(doc.part.partname),
                 "target_part": None if rel.is_external else str(rel.target_part.partname),
                  "external": bool(rel.is_external)}
        placed = by_relationship.get(rel.rId)
        if placed:
            placements_for_relationship = [{"position": item["position"], "placement": item["placement"]}
                                           for item in placed]
            for item in placements_for_relationship:
                images.append({**image, **item, "placements": placements_for_relationship})
        else:
            images.append(image)
    # OCR is strictly opt-in and its result remains separate from document text.
    ocr_results: list[dict[str, Any]] = []
    if ocr is not None:
        for rel in doc.part.rels.values():
            if "image" not in rel.reltype or rel.is_external:
                continue  # Never fetch externally linked image content for OCR.
            source = f"{path}#relationship={rel.rId}"
            result = ocr.recognize(rel.target_part.blob, source=source)
            if result.source != source or (result.confidence is not None and not 0 <= result.confidence <= 1):
                raise ValueError("OCR provider must return correct source and confidence between 0 and 1")
            ocr_results.append(asdict(result))
    # Inline and anchored drawings are counted separately from the document XML.
    xml = doc.element.xml
    placements = {"inline": len(re.findall(r"<wp:inline\b", xml)), "anchor": len(re.findall(r"<wp:anchor\b", xml))}
    return {"classification": c.to_dict(), "readable": True, "paragraphs": paragraphs, "tables": tables,
            "headers_footers": sections, "styles": styles, "comments": comments,
            "images": images, "image_placements": placements, "ocr": ocr_results,
            "text": "\n".join(p["text"] for p in paragraphs)}


def _clip_structure_text(text: str) -> tuple[str, bool]:
    """One bounded structure string plus whether it was shortened."""
    return (text[:_MAX_STRUCTURE_TEXT], True) if len(text) > _MAX_STRUCTURE_TEXT else (text, False)


def _sdt_value(properties: Any, name: str) -> str | None:
    """A content-control property such as w:tag, whose value lives in the w:val attribute."""
    if properties is None:
        return None
    child = properties.find("{%s}%s" % (_WORD_NS, name))
    if child is None:
        return None
    value = child.get("{%s}val" % _WORD_NS)
    if value is None:
        return None
    return value[:_MAX_STRUCTURE_TEXT]


def _revision_entry(node: Any, revision: str) -> dict[str, Any]:
    """One tracked insertion or deletion exactly as stored; nothing is applied."""
    w = "{%s}" % _WORD_NS
    text = "".join(part.text or "" for part in node.iter(w + "t"))
    if revision == "deletion":
        text += "".join(part.text or "" for part in node.iter(w + "delText"))
    clipped, was_clipped = _clip_structure_text(text)
    return {"revision": revision, "id": node.get(w + "id"), "author": node.get(w + "author"),
            "date": node.get(w + "date"), "text": clipped, "text_truncated": was_clipped}


def _content_control_entry(node: Any) -> dict[str, Any]:
    """One w:sdt block/run/cell/row control with its tag, alias and type marker."""
    w = "{%s}" % _WORD_NS
    properties = node.find(w + "sdtPr")
    content = node.find(w + "sdtContent")
    kinds = {child.tag.rsplit("}", 1)[-1] for child in (content if content is not None else [])}
    level = ("block" if "p" in kinds else "cell" if "tc" in kinds else
             "row" if "tr" in kinds else "run")
    text = "".join(part.text or "" for part in node.iter(w + "t"))
    clipped, was_clipped = _clip_structure_text(text)
    return {"level": level, "tag": _sdt_value(properties, "tag"),
            "alias": _sdt_value(properties, "alias"),
            "control_type": next((kind for kind in _SDT_KINDS
                                  if properties is not None and properties.find(w + kind) is not None),
                                 "unknown"),
            "text": clipped, "text_truncated": was_clipped}


def _word_revision_report(archive: zipfile.ZipFile, *, limit: int) -> dict[str, Any]:
    """Tracked changes and content controls, bounded and fail-closed.

    Word stores revisions (``w:ins``/``w:del``) and content controls (``w:sdt``) in the
    body, header and footer parts. Nothing here is evaluated or applied: items are
    reported as stored, every count is reported alongside the bounded list so a truncated
    answer cannot look complete, and a part that cannot be parsed refuses the read rather
    than reporting "no revisions". Parts skipped for size are listed as unscanned.
    """
    namespace = "{%s}" % _WORD_NS
    names = archive.namelist()
    parts = (["word/document.xml"] + sorted(
        name for name in names if re.fullmatch(r"word/(?:header|footer)\d*\.xml", name))
    )
    insertions: list[dict[str, Any]] = []
    deletions: list[dict[str, Any]] = []
    controls: list[dict[str, Any]] = []
    details: list[dict[str, Any]] = []
    insertion_count = deletion_count = move_count = control_count = detail_count = 0
    scanned: list[str] = []
    unscanned: list[str] = list(parts[_MAX_REVISION_PARTS:])
    for part in parts[:_MAX_REVISION_PARTS]:
        if part not in names:
            continue
        if archive.getinfo(part).file_size > _MAX_STRUCTURE_PART_BYTES:
            unscanned.append(part)
            continue
        try:
            root = ET.fromstring(archive.read(part))
        except ET.ParseError as exc:
            raise UnsupportedDocumentError(f"{part} could not be parsed safely") from exc
        scanned.append(part)
        parents = ({child: parent for parent in root.iter() for child in parent}
                   if len(details) < limit else {})
        part_index = 0
        for node in root.iter():
            if not node.tag.startswith(namespace):
                continue
            tag = node.tag[len(namespace):]
            if tag not in _WORD_REVISION_DETAIL_MARKS:
                continue
            detail_count += 1
            part_index += 1
            if len(details) < limit:
                label = "deletion" if tag in {"del", "moveFrom"} else tag
                entry = _revision_entry(node, label)
                entry["revision"] = tag
                parent = parents.get(node)
                details.append({**entry, "part": part, "part_index": part_index,
                                "tag": tag, "parent_tag": (parent.tag.rsplit("}", 1)[-1]
                                                          if parent is not None else None)})
        for node in root.iter(namespace + "ins"):
            insertion_count += 1
            if len(insertions) < limit:
                insertions.append(_revision_entry(node, "insertion"))
        for node in root.iter(namespace + "del"):
            deletion_count += 1
            if len(deletions) < limit:
                deletions.append(_revision_entry(node, "deletion"))
        move_count += sum(1 for _ in root.iter(namespace + "moveFrom"))
        move_count += sum(1 for _ in root.iter(namespace + "moveTo"))
        for node in root.iter(namespace + "sdt"):
            control_count += 1
            if len(controls) < limit:
                controls.append(_content_control_entry(node))
    return {"tracked_changes": {
                "insertions": insertions, "deletions": deletions,
                "insertion_count": insertion_count, "deletion_count": deletion_count,
                "move_mark_count": move_count,
                "insertions_truncated": insertion_count > len(insertions),
                "deletions_truncated": deletion_count > len(deletions),
                "revision_details": details, "revision_detail_count": detail_count,
                "revision_details_truncated": detail_count > len(details),
                "parts_scanned": scanned, "parts_unscanned": unscanned[:_MAX_REVISION_PARTS],
                "part_count": len(parts), "unscanned_part_count": len(unscanned),
                "parts_unscanned_truncated": len(unscanned) > _MAX_REVISION_PARTS,
                "counts_complete": not unscanned, "complete": not unscanned},
            "content_controls": controls, "content_control_count": control_count,
            "content_controls_truncated": control_count > len(controls)}


def read_word_structure(path: str | os.PathLike[str], *, limit: int = 20) -> dict[str, Any]:
    """Bounded read of DOCX headers/footers, comments, image anchors, revisions and controls.

    Body text is deliberately not materialised here: a large document would otherwise
    be loaded in full just to report structure. Use ``read_word_region`` for body text.
    Tracked changes and content controls are reported as stored; nothing is accepted or
    applied, and both lists carry the full count found next to the bounded list.
    """
    if not 1 <= limit <= 50:
        raise ValueError("limit must be between 1 and 50")
    c = classify_document(path)
    if c.kind != "word" or c.protected or not c.readable:
        raise UnsupportedDocumentError("only readable DOCX structure can be read")
    if _DocxDocument is None:
        raise UnsupportedDocumentError("python-docx is not installed")
    doc = _DocxDocument(path)

    def clipped(text: str) -> tuple[str, bool]:
        return (text[:_MAX_STRUCTURE_TEXT], True) if len(text) > _MAX_STRUCTURE_TEXT else (text, False)

    sections, sections_truncated = [], False
    for section in doc.sections:
        if len(sections) >= limit:
            sections_truncated = True
            break
        entry = {"header": [], "footer": []}
        for name, part in (("header", section.header), ("footer", section.footer)):
            for paragraph in part.paragraphs:
                if len(entry[name]) >= limit:
                    entry[f"{name}_truncated"] = True
                    break
                text, was_clipped = clipped(paragraph.text)
                entry[name].append({"text": text, "text_truncated": was_clipped,
                                    "style": paragraph.style.name if paragraph.style else None})
        sections.append(entry)

    comments, comments_truncated = [], False
    try:
        with zipfile.ZipFile(path) as archive:
            if "word/comments.xml" in archive.namelist():
                ns = {"w": _WORD_NS}
                for node in ET.fromstring(archive.read("word/comments.xml")).findall("w:comment", ns):
                    if len(comments) >= limit:
                        comments_truncated = True
                        break
                    text, was_clipped = clipped("".join(node.itertext()))
                    comments.append({"id": node.get("{%s}id" % _WORD_NS),
                                     "author": node.get("{%s}author" % _WORD_NS),
                                     "text": text, "text_truncated": was_clipped})
            # Revisions and content controls are read from the same package in one pass.
            revisions = _word_revision_report(archive, limit=limit)
    except (ET.ParseError, zipfile.BadZipFile):
        raise UnsupportedDocumentError("comments part could not be parsed safely")

    drawings = _docx_image_drawings(doc)
    images, images_truncated = [], False
    for item in drawings:
        if len(images) >= limit:
            images_truncated = True
            break
        images.append({**item, "source_part": "word/document.xml"})
    return {"sections": sections, "sections_truncated": sections_truncated,
            "comments": comments, "comments_truncated": comments_truncated,
            "images": images, "images_truncated": images_truncated,
            "image_placements": {"inline": sum(1 for item in drawings if item["placement"] == "inline"),
                                 "anchor": sum(1 for item in drawings if item["placement"] == "anchored")},
            "tracked_changes": revisions["tracked_changes"],
            "content_controls": revisions["content_controls"],
            "content_control_count": revisions["content_control_count"],
            "content_controls_truncated": revisions["content_controls_truncated"],
            "limit": limit,
            "note": "正文请用 read_word_region；此处只返回页眉页脚、批注与图片位置，"
                    "并按原文报告修订标记（w:ins/w:del，未接受也未拒绝）与内容控件（w:sdt）；"
                    "被 limit 截断的列表同时给出总数与 *_truncated 标志，未扫描的部件列在 "
                    "tracked_changes.parts_unscanned 中，不能用空列表当作“没有修订”。"}


def _table_tag(name: str) -> str:
    return "{%s}%s" % (_WORD_NS, name)


def _table_int(element: Any, name: str, default: int | None) -> int | None:
    """One integer *attribute*, or ``None`` when present but not a usable integer."""
    if element is None:
        return default
    raw = element.get(_table_tag(name))
    if raw is None:
        return default
    try:
        return int(raw)
    except (TypeError, ValueError):
        return None


def _table_val_int(parent: Any, name: str, default: int | None) -> int | None:
    """One integer ``w:val`` of a *child element* such as ``w:tcPr/w:gridSpan``."""
    if parent is None:
        return default
    element = parent.find(_table_tag(name))
    if element is None:
        return default
    raw = element.get(_table_tag("val"))
    if raw is None:
        return default
    try:
        return int(raw)
    except (TypeError, ValueError):
        return None


def _cell_paragraph_text(tc: Any) -> str:
    """Cell paragraph text as Word shows it: direct paragraphs, inserted runs included.

    Text inside a nested table is not part of the cell's own paragraph text, and text
    inside ``w:del`` is revision content rather than current content; both are reported
    separately instead of being merged into this string.
    """
    return "\n".join("".join(node.text or "" for node in paragraph.iter(_table_tag("t")))
                     for paragraph in tc.findall(_table_tag("p")))


def _cell_covering(rows: list[dict[str, Any]], row_index: int, column: int) -> dict[str, Any] | None:
    """The cell of one row whose horizontal extent contains a logical grid column."""
    if not 0 <= row_index < len(rows):
        return None
    for cell in rows[row_index]["cells"]:
        if cell["column"] <= column < cell["column"] + cell["grid_span"]:
            return cell
    return None


def _word_table_facts(tbl: Any) -> dict[str, Any]:
    """Per-cell merge facts for one ``w:tbl``, read from XML instead of python-docx.

    python-docx deliberately collapses a merged region into a single ``_Cell`` object,
    which is exactly the information a merge report must not lose. Reading ``w:tc``
    directly keeps the origin, its gridSpan followers and the vMerge continuations
    separate, gives physical and logical coordinates side by side, and records every
    inconsistency instead of smoothing it over.
    """
    issues: list[str] = []
    grid = tbl.find(_table_tag("tblGrid"))
    grid_columns = len(grid.findall(_table_tag("gridCol"))) if grid is not None else 0
    if grid_columns == 0:
        issues.append("表格缺少 w:tblGrid，逻辑列数无法确定")
    rows: list[dict[str, Any]] = []
    for row_index, tr in enumerate(tbl.findall(_table_tag("tr"))):
        tr_pr = tr.find(_table_tag("trPr"))
        grid_before = _table_int(tr_pr, "gridBefore", 0)
        grid_after = _table_int(tr_pr, "gridAfter", 0)
        if grid_before is None or grid_after is None:
            issues.append(f"第 {row_index} 行的 w:gridBefore/w:gridAfter 不是整数，按 0 处理")
        grid_before = grid_before if isinstance(grid_before, int) and grid_before >= 0 else 0
        grid_after = grid_after if isinstance(grid_after, int) and grid_after >= 0 else 0
        cells: list[dict[str, Any]] = []
        column = grid_before
        for physical_index, tc in enumerate(tr.findall(_table_tag("tc"))):
            tc_pr = tc.find(_table_tag("tcPr"))
            span = _table_val_int(tc_pr, "gridSpan", 1)
            if not isinstance(span, int) or span < 1:
                issues.append(f"第 {row_index} 行第 {physical_index} 个 w:tc 的 w:gridSpan "
                              f"不是正整数，按 1 处理")
                span = 1
            vmerge_element = tc_pr.find(_table_tag("vMerge")) if tc_pr is not None else None
            vmerge = None
            if vmerge_element is not None:
                value = (vmerge_element.get(_table_tag("val")) or "").strip().lower()
                if value in ("", "continue"):
                    vmerge = "continue"
                elif value == "restart":
                    vmerge = "restart"
                else:
                    issues.append(f"第 {row_index} 行第 {physical_index} 个 w:tc 的 w:vMerge "
                                  f"取值 {value!r} 无法识别")
                    vmerge = "continue"
            legacy = tc_pr is not None and tc_pr.find(_table_tag("hMerge")) is not None
            if legacy:
                issues.append(f"第 {row_index} 行第 {physical_index} 个 w:tc 使用旧式 w:hMerge，"
                              f"本工具不解释这种合并")
            text, was_clipped = _clip_structure_text(_cell_paragraph_text(tc))
            cells.append({"row": row_index, "column": column, "last_column": column + span - 1,
                          "physical_index": physical_index, "grid_span": span, "vmerge": vmerge,
                          "legacy_hmerge": legacy, "text": text, "text_truncated": was_clipped,
                          "nested_tables": list(tc.findall(_table_tag("tbl"))),
                          "element": tc, "anchor": {"row": row_index, "column": column}})
            column += span
        covered = column + grid_after
        if grid_columns and covered != grid_columns:
            issues.append(f"第 {row_index} 行覆盖 {covered} 列，与 w:tblGrid 的 {grid_columns} 列不一致")
        rows.append({"index": row_index, "physical_cells": len(cells), "grid_before": grid_before,
                     "grid_after": grid_after, "grid_covered": covered, "cells": cells})
    # A vMerge continuation must resolve upward to a restart; otherwise the file claims a
    # merged region that does not exist, and the report says so instead of inventing one.
    for row_index, row in enumerate(rows):
        for cell in row["cells"]:
            if cell["vmerge"] != "continue":
                continue
            above = _cell_covering(rows, row_index - 1, cell["column"])
            if above is None or above["vmerge"] is None:
                issues.append(f"第 {row_index} 行第 {cell['column']} 列的 w:vMerge 延续格在上方找不到起点")
                continue
            if above["grid_span"] != cell["grid_span"]:
                issues.append(f"第 {row_index} 行第 {cell['column']} 列的合并宽度 {cell['grid_span']} "
                              f"与上方起点 {above['grid_span']} 不一致")
            cell["anchor"] = dict(above["anchor"])
    return {"grid_columns": grid_columns, "rows": rows, "issues": issues}


def _nested_table_report(facts: dict[str, Any], *, limit: int, depth: int,
                         path: list[dict[str, int]],
                         found: list[dict[str, Any]]) -> tuple[bool, bool]:
    """Nested tables below one table. Returns ``(count_limit_reached, depth_limit_reached)``."""
    for row in facts["rows"]:
        for cell in row["cells"]:
            for nested in cell["nested_tables"]:
                if len(found) >= limit:
                    return True, False
                entry_path = path + [{"row": cell["row"], "column": cell["column"]}]
                nested_facts = _word_table_facts(nested)
                found.append({"depth": depth, "path": entry_path,
                              "rows": len(nested_facts["rows"]),
                              "grid_columns": nested_facts["grid_columns"],
                              "merge_consistent": not nested_facts["issues"]})
                if depth >= _MAX_TABLE_DEPTH:
                    # Deeper structure is deliberately not expanded; report that it exists.
                    if any(len(inner["nested_tables"]) for row_facts in nested_facts["rows"]
                           for inner in row_facts["cells"]):
                        return False, True
                    continue
                count_cut, depth_cut = _nested_table_report(nested_facts, limit=limit, depth=depth + 1,
                                                            path=entry_path, found=found)
                if depth_cut or count_cut:
                    return count_cut, depth_cut
    return False, False


def word_table_grid(tbl: Any, *, limit: int = 50) -> dict[str, Any]:
    """Bounded, checkable view of one ``w:tbl``: grid, merges, nesting and consistency."""
    facts = _word_table_facts(tbl)
    rows = []
    for row in facts["rows"][:limit]:
        rows.append({"index": row["index"], "physical_cells": row["physical_cells"],
                     "grid_before": row["grid_before"], "grid_after": row["grid_after"],
                     "grid_covered": row["grid_covered"],
                     "cells": [{**{key: value for key, value in cell.items()
                                   if key not in {"element", "nested_tables"}},
                                "nested_table_count": len(cell["nested_tables"])}
                               for cell in row["cells"]]})
    continuations = [{"row": cell["row"], "column": cell["column"], "anchor": dict(cell["anchor"]),
                      "grid_span": cell["grid_span"]}
                     for row in facts["rows"] for cell in row["cells"]
                     if (cell["anchor"]["row"], cell["anchor"]["column"]) != (cell["row"], cell["column"])]
    nested: list[dict[str, Any]] = []
    nested_cut, depth_cut = _nested_table_report(facts, limit=limit, depth=1, path=[], found=nested)
    issues = facts["issues"]
    return {"grid_columns": facts["grid_columns"], "row_count": len(facts["rows"]), "rows": rows,
            "cells_truncated": len(facts["rows"]) > limit,
            "continuations": continuations[:limit],
            "continuations_truncated": len(continuations) > limit,
            "nested_tables": nested, "nested_tables_truncated": nested_cut,
            "depth_limit_reached": depth_cut, "max_nested_depth": _MAX_TABLE_DEPTH,
            "merge_consistent": not issues, "merge_issues": issues[:_MAX_TABLE_ISSUES],
            "merge_issue_count": len(issues), "merge_issues_truncated": len(issues) > _MAX_TABLE_ISSUES,
            "limit": limit}


def word_table_merge_issues(tbl: Any) -> list[str]:
    """Structural merge problems that make a table unsafe to rewrite, if any."""
    return list(_word_table_facts(tbl)["issues"])


def locate_word_cell(tbl: Any, row: int, column: int) -> tuple[Any, dict[str, Any]]:
    """The ``w:tc`` carrying the content of logical cell ``(row, column)``, plus its facts.

    Only the cell that owns the content can be addressed: a gridSpan follower or a vMerge
    continuation must be written through its origin, so a refusal names that origin rather
    than silently writing into a cell whose text Word does not display there.
    """
    facts = _word_table_facts(tbl)
    if not 0 <= row < len(facts["rows"]):
        raise ValueError(f"行 {row} 超出表格范围：该表共 {len(facts['rows'])} 行（0 起）")
    row_facts = facts["rows"][row]
    cell = next((item for item in row_facts["cells"]
                 if item["column"] <= column < item["column"] + item["grid_span"]), None)
    if cell is None:
        raise ValueError(f"列 {column} 超出表格范围：第 {row} 行覆盖 {row_facts['grid_covered']} 列")
    if cell["column"] != column:
        raise ValueError(f"逻辑坐标 (行 {row}, 列 {column}) 被 (行 {row}, 列 {cell['column']}) 的跨列合并覆盖"
                         f"（gridSpan={cell['grid_span']}）；该单元格的文本由起点承载，请改向起点写入")
    if cell["vmerge"] == "continue":
        anchor = cell["anchor"]
        raise ValueError(f"逻辑坐标 (行 {row}, 列 {column}) 是纵向合并的延续格（vMerge=continue），"
                         f"文本由 (行 {anchor['row']}, 列 {anchor['column']}) 承载，请改向该起点写入")
    return cell["element"], {key: value for key, value in cell.items()
                             if key not in {"element", "nested_tables"}}


def read_word_table_structure(path: str | os.PathLike[str], *, table_index: int = 0,
                              limit: int = 50) -> dict[str, Any]:
    """Bounded, checkable structure of one DOCX table: merges, nesting and consistency.

    Every cell is reported with its physical index and logical grid column, its gridSpan
    and vMerge state, and the origin cell that carries its content. Nested tables are
    listed with their path and depth (bounded, with ``depth_limit_reached`` when the cap
    is hit). Structural inconsistencies — row coverage that disagrees with ``w:tblGrid``,
    an orphan continuation, an unbalanced span, a legacy ``w:hMerge`` — are reported as
    issues with ``merge_consistent=false`` rather than being smoothed over.
    """
    if not 1 <= limit <= _MAX_STRUCTURE_LIMIT:
        raise ValueError(f"limit must be between 1 and {_MAX_STRUCTURE_LIMIT}")
    if table_index < 0:
        raise ValueError("table_index must not be negative")
    c = classify_document(path)
    if c.kind != "word" or c.protected or not c.readable:
        raise UnsupportedDocumentError("only readable DOCX table structure can be read")
    if _DocxDocument is None:
        raise UnsupportedDocumentError("python-docx is not installed")
    doc = _DocxDocument(path)
    if table_index >= len(doc.tables):
        raise ValueError(f"table index {table_index} is out of bounds for "
                         f"{len(doc.tables)} top-level table(s)")
    report = word_table_grid(doc.tables[table_index]._tbl, limit=limit)
    report.update({"table_index": table_index, "table_count": len(doc.tables), "readable": True,
                   "note": "column 是逻辑网格列（0 起），physical_index 是 w:tr 内 w:tc 的物理序号；"
                           "只有 anchor 指向的单元格承载文本，被 gridSpan 覆盖或 vMerge=continue 的坐标"
                           "必须改用 anchor；嵌套表格按 depth 列出且不超过 max_nested_depth，"
                           "触顶时 depth_limit_reached 为真；limit 截断时对应 *_truncated 为真，"
                           "merge_issues 只列前若干条而 merge_issue_count 是总数；"
                           "table_index 只索引正文顶层表格。"})
    return report


def read_word_text(path: str | os.PathLike[str], *, offset: int = 0, limit: int = 20,
                   max_chars: int = 4000) -> dict[str, Any]:
    """Read a DOCX in document order, covering body paragraphs *and* table cells.

    A request to "show the text" must not silently stop at the body: documents whose
    content lives in a table have almost no paragraph text, so paragraph-only reading
    returns nearly nothing. Blocks are returned in true document order and paginated, so
    a large file stays inside the decision budget and the context limit.
    """
    if offset < 0:
        raise ValueError("offset must not be negative")
    if not 1 <= limit <= 50:
        raise ValueError("limit must be between 1 and 50")
    if not 200 <= max_chars <= 20000:
        raise ValueError("max_chars must be between 200 and 20000")
    # A read smaller than this cannot show anything useful, and small models routinely send
    # the minimum allowed value. Raising the floor is safe: the model's own context only
    # ever holds a bounded excerpt of a result, so this affects what the *user* sees, not
    # how much is pushed into the model. Both requested and effective values are reported.
    effective_limit = max(limit, _MIN_TEXT_BLOCKS)
    effective_chars = max(max_chars, _MIN_TEXT_BUDGET)
    c = classify_document(path)
    if c.kind != "word" or c.protected or not c.readable:
        raise UnsupportedDocumentError("only readable DOCX text can be read")
    if _DocxDocument is None:
        raise UnsupportedDocumentError("python-docx is not installed")
    doc = _DocxDocument(path)

    # ``max_chars`` bounds the *text* in one reply; ``limit`` bounds block count. A block
    # is never skipped for being large, and a table never silently loses rows: what was
    # left out is flagged, together with the tool that can fetch the remainder. The cursor
    # always advances past a returned block, so paging cannot spin.
    budget = effective_chars
    kept: list[dict[str, Any]] = []
    for index, block in enumerate(doc.iter_inner_content()):
        if index < offset:
            continue
        if len(kept) >= effective_limit or budget <= 0:
            break
        if isinstance(block, _DocxTable):
            rows: list[list[str]] = []
            omitted_rows = 0
            for row in block.rows:
                if budget <= 0 or len(rows) >= effective_limit:
                    omitted_rows += 1
                    continue
                texts = []
                for cell in row.cells:
                    take = cell.text[:budget]
                    budget -= len(take)
                    texts.append(take)
                    if budget <= 0:
                        break
                rows.append(texts)
            entry: dict[str, Any] = {"kind": "table", "index": index, "rows": rows,
                                     "row_count": len(block.rows), "column_count": len(block.columns),
                                     "rows_truncated": omitted_rows > 0}
            if omitted_rows:
                entry["rows_omitted"] = omitted_rows
                entry["remainder_hint"] = ("其余行请用 read_word_table_region 指定 start_row 继续读取，"
                                           "或提高 max_chars 后重试")
            kept.append(entry)
        else:
            text = block.text
            take = text[:budget]
            budget -= len(take)
            entry = {"kind": "paragraph", "index": index, "text": take,
                     "text_truncated": take != text,
                     "style": block.style.name if block.style else None}
            if entry["text_truncated"]:
                entry["remainder_hint"] = "该段更长的原文可用 read_word_region 按段索引读取"
            kept.append(entry)
    consumed = offset + len(kept)
    total = _count_word_blocks(doc)
    # Anything held back for space must keep has_more true, or a caller would believe it
    # had read the whole document.
    truncated = any(block.get("rows_truncated") or block.get("text_truncated") for block in kept)
    return {"blocks": kept, "offset": offset, "next_offset": consumed,
            "limit": limit, "effective_limit": effective_limit,
            "max_chars": max_chars, "effective_max_chars": effective_chars,
            "returned_chars": effective_chars - max(0, budget),
            "total_blocks": total, "has_more": consumed < total or truncated,
            "truncated_in_this_reply": truncated,
            "note": "按文档顺序返回正文段落与表格单元格（表格内容也在其中）；"
                    "若 has_more 为真，先看各块的 remainder_hint：未读完的块用提示的工具续读，"
                    "其余块用 next_offset 继续。"}


def _count_word_blocks(doc: Any) -> int:
    return sum(1 for _ in doc.iter_inner_content())


def summarize_word_document(path: str | os.PathLike[str], *, chunk_size: int = 20,
                            max_chunks: int = 25, preview_chars: int = 200) -> dict[str, Any]:
    """Bounded outline plus chunk map of a DOCX, so a summary can be planned in a few steps.

    Reading every paragraph through ``read_word_region`` would need one model decision per
    20 paragraphs, which cannot fit the 15-decision budget on a real document. This returns
    the heading outline, per-chunk character counts and a short preview of each chunk; the
    model then pulls only the chunks it needs and writes the summary itself.
    """
    if not 1 <= chunk_size <= 200:
        raise ValueError("chunk_size must be between 1 and 200")
    if not 1 <= max_chunks <= 100:
        raise ValueError("max_chunks must be between 1 and 100")
    if not 1 <= preview_chars <= 1000:
        raise ValueError("preview_chars must be between 1 and 1000")
    c = classify_document(path)
    if c.kind != "word" or c.protected or not c.readable:
        raise UnsupportedDocumentError("only readable DOCX can be summarised")
    if _DocxDocument is None:
        raise UnsupportedDocumentError("python-docx is not installed")
    doc = _DocxDocument(path)

    headings: list[dict[str, Any]] = []
    headings_truncated = False
    chunks: list[dict[str, Any]] = []
    characters = 0
    paragraphs = doc.paragraphs
    for index, paragraph in enumerate(paragraphs):
        text = paragraph.text
        characters += len(text)
        if paragraph.style and paragraph.style.name.lower().startswith("heading"):
            if len(headings) >= _MAX_OUTLINE_HEADINGS:
                headings_truncated = True
            else:
                headings.append({"index": index, "style": paragraph.style.name,
                                 "text": text[:_MAX_STRUCTURE_TEXT],
                                 "text_truncated": len(text) > _MAX_STRUCTURE_TEXT})
        if index % chunk_size == 0:
            if len(chunks) >= max_chunks:
                continue
            chunks.append({"start_index": index, "end_index": min(index + chunk_size, len(paragraphs)) - 1,
                           "characters": 0, "preview": "", "preview_truncated": False})
        if chunks:
            current = chunks[-1]
            if current["start_index"] <= index <= current["end_index"]:
                current["characters"] += len(text)
                if len(current["preview"]) < preview_chars:
                    current["preview"] = (current["preview"] + " " + text).strip()[:preview_chars]
                    current["preview_truncated"] = len(current["preview"]) >= preview_chars
    total_chunks = (len(paragraphs) + chunk_size - 1) // chunk_size if paragraphs else 0
    # Report table shapes so a request for "the text" does not stop at the body: a document
    # whose content sits in a table looks empty through read_word_region alone.
    tables = [{"index": index, "rows": len(table.rows), "columns": len(table.columns)}
              for index, table in enumerate(doc.tables)][:_MAX_TABLE_SUMMARIES]
    return {"paragraph_count": len(paragraphs), "table_count": len(doc.tables),
            "tables": tables,
            "tables_truncated": len(doc.tables) > len(tables),
            "character_count": characters, "headings": headings,
            "headings_truncated": headings_truncated, "chunk_size": chunk_size,
            "chunks": chunks, "total_chunks": total_chunks,
            "chunks_omitted": max(0, total_chunks - len(chunks)),
            "note": "先用本大纲确定要读的分块，再用 read_word_region 取正文；"
                    "表格内容不在正文里，本字段列出每张表的行/列数，请用 read_word_table_region 读取；"
                    "摘要由模型撰写，本工具只做有界检索。"}


def search_docx(path: str | os.PathLike[str], query: str, *, chunk_size: int = 1000, overlap: int = 100,
                case_sensitive: bool = False, max_results: int = 50) -> list[dict[str, Any]]:
    """Search DOCX paragraphs and table cells without building a full report."""
    c = classify_document(path)
    if c.kind != "word" or c.protected or not c.readable or not query or max_results <= 0:
        return []
    if _DocxDocument is None:
        raise UnsupportedDocumentError("python-docx is not installed")
    if chunk_size <= 0:
        raise ValueError("chunk_size must be positive")
    doc = _DocxDocument(path)
    needle = query if case_sensitive else query.lower()
    found: list[dict[str, Any]] = []

    def add_matches(text: str, location: dict[str, Any]) -> None:
        if not text:
            return
        hay = text if case_sensitive else text.lower()
        start = 0
        while len(found) < max_results:
            match = hay.find(needle, start)
            if match < 0:
                break
            left = max(0, match - max(0, chunk_size - len(query)) // 2)
            snippet = text[left:left + chunk_size]
            found.append({"text": snippet, "location": location, "offset": match,
                          "match": text[match:match + len(query)]})
            start = match + max(len(query), 1)

    for index, paragraph in enumerate(doc.paragraphs):
        add_matches(paragraph.text, {"type": "paragraph", "index": index})
        if len(found) >= max_results:
            return found
    for table_index, table in enumerate(doc.tables):
        for row_index, row in enumerate(table.rows):
            for column_index, cell in enumerate(row.cells):
                add_matches(cell.text, {"type": "table_cell", "table": table_index,
                                        "row": row_index, "column": column_index})
                if len(found) >= max_results:
                    return found
    return found


def inspect_xlsx(path: str | os.PathLike[str]) -> dict[str, Any]:
    c = classify_document(path)
    if c.kind != "spreadsheet" or c.protected or not c.readable:
        return {"classification": c.to_dict(), "readable": False, "limitations": list(c.limitations)}
    if openpyxl is None: raise UnsupportedDocumentError("openpyxl is not installed")
    images_by_sheet = _xlsx_sheet_images(path)
    formulas = openpyxl.load_workbook(path, data_only=False, read_only=False, keep_vba=c.macro_enabled)
    try:
        cached = openpyxl.load_workbook(path, data_only=True, read_only=False, keep_vba=c.macro_enabled)
        try:
            return _inspect_loaded_xlsx(formulas, cached, c, images_by_sheet)
        finally:
            cached.close()
    finally:
        formulas.close()


_MAIN_NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
# Relationship *type* strings carry a trailing slash; the *attribute* namespace does not.
_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_REL_TYPE = _REL_NS + "/"
_DRAWING_NS = {"xdr": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
               "a": "http://schemas.openxmlformats.org/drawingml/2006/main"}
_MAX_DRAWINGS = 200


def _part_rels(archive: zipfile.ZipFile, part: str) -> dict[str, tuple[str, str]]:
    """``{relationship id: (type, target)}`` for one package part, or {} if absent."""
    directory, _, name = part.rpartition("/")
    rels_path = f"{directory}/_rels/{name}.rels"
    if rels_path not in archive.namelist():
        return {}
    root = ET.fromstring(archive.read(rels_path))
    return {rel.get("Id", ""): (rel.get("Type", ""), rel.get("Target", ""))
            for rel in root if rel.tag.endswith("Relationship")}


def _resolve_part(base_part: str, target: str) -> str:
    """Resolve an OOXML relationship target, which may be relative or package-absolute."""
    combined = target.lstrip("/") if target.startswith("/") else base_part.rsplit("/", 1)[0] + "/" + target
    pieces: list[str] = []
    for piece in combined.split("/"):
        if piece in ("", "."):
            continue
        if piece == "..":
            if pieces:
                pieces.pop()
            continue
        pieces.append(piece)
    return "/".join(pieces)


def _drawing_pictures(archive: zipfile.ZipFile, drawing_part: str) -> list[dict[str, Any]]:
    """Picture anchors inside one drawing part, without decoding any image bytes."""
    rels = _part_rels(archive, drawing_part)
    pictures = []
    for anchor in ET.fromstring(archive.read(drawing_part)):
        anchor_kind = anchor.tag.rsplit("}", 1)[-1]
        if anchor_kind not in {"oneCellAnchor", "twoCellAnchor"} or anchor.find("xdr:pic", _DRAWING_NS) is None:
            continue
        start = anchor.find("xdr:from", _DRAWING_NS)
        position = None
        if start is not None:
            column = (start.findtext("xdr:col", default="", namespaces=_DRAWING_NS) or "").strip()
            row = (start.findtext("xdr:row", default="", namespaces=_DRAWING_NS) or "").strip()
            if column.isdigit() and row.isdigit():
                position = {"type": "cell",
                            "coordinate": f"{get_column_letter(int(column) + 1)}{int(row) + 1}",
                            "row": int(row) + 1, "column": int(column) + 1}
        blip = anchor.find(".//a:blip", _DRAWING_NS)
        relationship = blip.get(f"{{{_REL_NS}}}embed") if blip is not None else None
        target = rels.get(relationship or "", ("", ""))[1]
        pictures.append({"type": "picture", "anchor": anchor_kind, "position": position,
                         "source_part": drawing_part,
                         "relationship": relationship,
                         "target_part": _resolve_part(drawing_part, target) if target else None})
        if len(pictures) >= _MAX_DRAWINGS:
            break
    return pictures


def _xlsx_drawing_report(path: str | os.PathLike[str]) -> tuple[dict[str, list[dict[str, Any]]], bool]:
    """Image anchors per worksheet, plus whether the drawing parts could be read at all.

    openpyxl's drawing reader returns no images at all when Pillow is missing, which
    would make an image-bearing workbook look image-free. The drawing XML is the
    authoritative source for image presence and anchor position; decoding the image
    bytes is separate and reported as such. The boolean keeps "no images" distinct from
    "the drawing parts could not be parsed", so an unreadable package is never reported
    as an image-free workbook.
    """
    by_sheet: dict[str, list[dict[str, Any]]] = {}
    try:
        with zipfile.ZipFile(path) as archive:
            names = set(archive.namelist())
            if "xl/workbook.xml" not in names:
                return {}, False
            workbook = ET.fromstring(archive.read("xl/workbook.xml"))
            workbook_rels = _part_rels(archive, "xl/workbook.xml")
            for sheet in workbook.iter(_MAIN_NS + "sheet"):
                title = sheet.get("name")
                relation = workbook_rels.get(sheet.get(f"{{{_REL_NS}}}id", ""), ("", ""))
                if not title or relation[0] != _REL_TYPE + "worksheet" or not relation[1]:
                    continue
                sheet_part = _resolve_part("xl/workbook.xml", relation[1])
                pictures: list[dict[str, Any]] = []
                for kind, target in _part_rels(archive, sheet_part).values():
                    if kind != _REL_TYPE + "drawing" or not target:
                        continue
                    drawing_part = _resolve_part(sheet_part, target)
                    if drawing_part in names:
                        pictures.extend(_drawing_pictures(archive, drawing_part))
                        pictures = pictures[:_MAX_DRAWINGS]
                if pictures:
                    by_sheet[title] = pictures
    except (zipfile.BadZipFile, ET.ParseError, OSError, KeyError):
        return {}, False  # Unparseable drawings must never be reported as absent images.
    return by_sheet, True


def _xlsx_sheet_images(path: str | os.PathLike[str]) -> dict[str, list[dict[str, Any]]]:
    return _xlsx_drawing_report(path)[0]


def _inspect_loaded_xlsx(formulas: Any, cached: Any, c: DocumentClassification,
                         images_by_sheet: dict[str, list[dict[str, Any]]] | None = None) -> dict[str, Any]:
    sheets = []
    for ws in formulas.worksheets:
        cw = cached[ws.title]
        cells = []
        populated = (cell for cell in ws._cells.values() if cell.value is not None)
        truncated = False
        for cell in populated:
            if len(cells) >= 5000:
                truncated = True
                break
            value = {"coordinate": cell.coordinate, "value": cell.value, "data_type": cell.data_type,
                     "style_id": cell.style_id, "number_format": cell.number_format}
            if cell.data_type == "f":
                value["formula_text"] = cell.value
                value["cached_value"] = cw[cell.coordinate].value
                value["cached_value_unverified"] = True
            cells.append(value)
        images = []
        for image in getattr(ws, "_images", []):
            anchor = getattr(image, "anchor", None)
            start = getattr(anchor, "_from", None)
            position = None
            if start is not None and hasattr(start, "row") and hasattr(start, "col"):
                position = {"type": "cell", "coordinate": ws.cell(start.row + 1, start.col + 1).coordinate,
                            "row": start.row + 1, "column": start.col + 1}
            images.append({"type": type(image).__name__, "position": position})
        # Package-derived anchors are authoritative: openpyxl drops every image when
        # Pillow is missing, so its empty list must not be read as "no images".
        from_package = list((images_by_sheet or {}).get(ws.title, [])) if images_by_sheet is not None else []
        if from_package:
            for index, entry in enumerate(from_package):
                entry["bytes_decoded"] = index < len(images)
            images = from_package
        sheets.append({"title": ws.title, "dimensions": ws.calculate_dimension(), "cells": cells,
                       "cells_truncated": truncated, "cell_preview_limit": 5000,
                       "merged_ranges": [str(r) for r in ws.merged_cells.ranges],
                       "charts": [{"type": type(chart).__name__, "title": str(chart.title) if chart.title else None}
                                  for chart in ws._charts],
                       "images": images, "max_row": ws.max_row, "max_column": ws.max_column})
    return {"classification": c.to_dict(), "readable": True, "sheets": sheets,
            "formula_values_are_distinct_from_cached_values": True,
            "cached_values_unverified": True}


_REFERENCE_PART = r"\$?[A-Za-z]{1,3}\$?[1-9][0-9]{0,6}"
_REFERENCE = re.compile(rf"(?:(?:'[^']+'|[A-Za-z_][A-Za-z0-9_.]*)!)?"
                        rf"{_REFERENCE_PART}(?::{_REFERENCE_PART})?")
_REFERENCE_UNION = re.compile(rf"(?:(?:'[^']+'|[A-Za-z_][A-Za-z0-9_.]*)!)?"
                              rf"{_REFERENCE_PART}(?::{_REFERENCE_PART})?(?:,(?:'[^']+'|[A-Za-z_][A-Za-z0-9_.]*)?"
                              rf"!?{_REFERENCE_PART}(?::{_REFERENCE_PART})?)*")


def _defined_name_kind(refers_to: str) -> str:
    """Classify a stored name target without evaluating anything.

    Only shapes that can be read literally are named: a single reference, a reference
    union, or a quoted constant. Everything else is reported as ``formula_or_unknown``
    rather than guessed at.
    """
    candidate = refers_to.strip()
    if candidate.startswith("="):
        candidate = candidate[1:].strip()
    if not candidate:
        return "empty"
    if len(candidate) >= 2 and candidate.startswith('"') and candidate.endswith('"'):
        return "constant"
    if _REFERENCE.fullmatch(candidate):
        return "range"
    if "," in candidate and _REFERENCE_UNION.fullmatch(candidate):
        return "range_union"
    return "formula_or_unknown"


def _compress_runs(indices: Iterable[int]) -> list[dict[str, Any]]:
    """Contiguous indices collapsed into runs, so one hidden block is one entry."""
    runs: list[dict[str, Any]] = []
    for index in sorted(set(indices)):
        if runs and index == runs[-1]["last"] + 1:
            runs[-1]["last"] = index
            runs[-1]["count"] += 1
            runs[-1]["label"] = f"{runs[-1]['first']}-{index}"
        else:
            runs.append({"first": index, "last": index, "count": 1, "label": str(index)})
    return runs


def _hidden_row_runs(ws: Any) -> tuple[list[dict[str, Any]], int]:
    """Hidden rows from row dimensions, including explicit zero height."""
    hidden: list[int] = []
    for index, dimension in ws.row_dimensions.items():
        if not isinstance(index, int) or not 1 <= index <= 1048576:
            continue
        if dimension.hidden or dimension.ht == 0:
            hidden.append(index)
    return _compress_runs(hidden), len(hidden)


def _hidden_column_runs(ws: Any) -> tuple[list[dict[str, Any]], int]:
    """Hidden columns from column dimensions, including explicit zero width."""
    hidden: set[int] = set()
    for key, dimension in ws.column_dimensions.items():
        if not (dimension.hidden or dimension.width == 0):
            continue
        try:
            low = dimension.min or (column_index_from_string(key) if column_index_from_string else 0)
        except (ValueError, TypeError):
            continue
        if not low:
            continue
        high = dimension.max or low
        for index in range(max(1, int(low)), min(int(high), 16384) + 1):
            hidden.add(index)
    runs = _compress_runs(hidden)
    for run in runs:
        run["label"] = (f"{get_column_letter(run['first'])}-{get_column_letter(run['last'])}"
                        if run["count"] > 1 else get_column_letter(run["first"]))
    return runs, len(hidden)


def _bounded_text(value: Any, limit: int = _MAX_NAMED_REFERENCE) -> tuple[str | None, bool]:
    if value is None:
        return None, False
    text = str(value)
    return (text[:limit], True) if len(text) > limit else (text, False)


def _chart_title_text(chart: Any) -> tuple[str | None, bool]:
    """Chart title text plus whether a title exists at all, without repr fallbacks."""
    title = getattr(chart, "title", None)
    if title is None:
        return None, False
    transaction = getattr(title, "tx", None)
    rich = getattr(transaction, "rich", None)
    if rich is not None:
        text = "".join(run.t for paragraph in (rich.p or []) for run in (paragraph.r or []) if run.t)
        if text:
            return text[:200], False
    reference = getattr(transaction, "strRef", None)
    formula = getattr(reference, "f", None) if reference is not None else None
    if formula:
        return str(formula)[:200], False
    return None, True  # A title exists but its text could not be read; never report "no title".


def _chart_anchor_position(chart: Any) -> dict[str, Any] | None:
    anchor = getattr(chart, "anchor", None)
    if anchor is None:
        return None
    start = getattr(anchor, "_from", None)
    if start is not None and hasattr(start, "row") and hasattr(start, "col"):
        return {"type": "cell", "coordinate": f"{get_column_letter(start.col + 1)}{start.row + 1}",
                "row": start.row + 1, "column": start.col + 1}
    return {"type": "unknown", "anchor_class": type(anchor).__name__}


def _chart_inventory(charts: Iterable[Any], *, limit: int) -> tuple[list[dict[str, Any]], int]:
    entries, total = [], 0
    for chart in charts:
        total += 1
        if len(entries) >= limit:
            continue
        title, title_unreadable = _chart_title_text(chart)
        series = []
        for item in list(getattr(chart, "series", []))[:limit]:
            values = (item.val.numRef.f if getattr(item, "val", None) is not None and item.val.numRef else
                      item.val.strRef.f if getattr(item, "val", None) is not None and item.val.strRef else None)
            categories = (item.cat.numRef.f if getattr(item, "cat", None) is not None and item.cat.numRef else
                          item.cat.strRef.f if getattr(item, "cat", None) is not None and item.cat.strRef else None)
            series.append({"values": _bounded_text(values)[0], "categories": _bounded_text(categories)[0]})
        entries.append({"type": type(chart).__name__, "chart_type": getattr(chart, "type", None),
                        "title": title, "title_unreadable": title_unreadable,
                        "anchor": _chart_anchor_position(chart),
                        "series_count": len(getattr(chart, "series", [])), "series": series})
    return entries, total


def _excel_sheet_structure(ws: Any, images: list[dict[str, Any]], *, limit: int) -> dict[str, Any]:
    hidden_rows, hidden_row_count = _hidden_row_runs(ws)
    hidden_columns, hidden_column_count = _hidden_column_runs(ws)
    formatting, rule_count, formatting_count = [], 0, 0
    for entry in ws.conditional_formatting:
        formatting_count += 1
        rules = []
        for rule in entry.rules:
            rule_count += 1
            formulas = []
            formula_truncated = False
            for formula in list(rule.formula or [])[:limit]:
                text, truncated = _bounded_text(formula)
                formulas.append(text)
                formula_truncated = formula_truncated or truncated
            rules.append({"type": rule.type, "operator": rule.operator, "priority": rule.priority,
                          "formula": formulas, "formula_truncated": formula_truncated,
                          "text": _bounded_text(rule.text)[0], "stop_if_true": rule.stopIfTrue,
                          "has_differential_format": rule.dxf is not None})
        if len(formatting) < limit:
            formatting.append({"range": str(entry.sqref), "rules": rules})
    charts, chart_count = _chart_inventory(getattr(ws, "_charts", []), limit=limit)
    merged = [str(rng) for rng in ws.merged_cells.ranges]
    return {"title": ws.title, "state": ws.sheet_state, "sheet_kind": "worksheet",
            "hidden_rows": hidden_rows[:limit], "hidden_row_count": hidden_row_count,
            "hidden_rows_truncated": len(hidden_rows) > limit,
            "hidden_columns": hidden_columns[:limit], "hidden_column_count": hidden_column_count,
            "hidden_columns_truncated": len(hidden_columns) > limit,
            "conditional_formatting": formatting, "conditional_formatting_rule_count": rule_count,
            "conditional_formatting_range_count": formatting_count,
            "conditional_formatting_truncated": formatting_count > len(formatting),
            "charts": charts, "chart_count": chart_count, "charts_truncated": chart_count > len(charts),
            "images": images[:limit], "image_count": len(images), "images_truncated": len(images) > limit,
            "merged_ranges": merged[:limit], "merged_range_count": len(merged),
            "merged_ranges_truncated": len(merged) > limit}


def read_excel_structure(path: str | os.PathLike[str], *, limit: int = 50) -> dict[str, Any]:
    """Bounded, read-only inventory of workbook semantics that cell reads cannot show.

    Named ranges, conditional formatting, hidden rows/columns, merged ranges, the chart
    inventory and package-level image anchors are reported exactly as stored. Nothing is
    evaluated: conditional-format rules are listed, never applied, and no formula is
    calculated. Every bounded list is accompanied by the full count found and a
    ``*_truncated`` flag, and unreadable drawing parts are reported as unreadable rather
    than as "no images". Hidden rows/columns come from the file's row/column dimensions,
    so view-only state (a collapsed outline group, a filtered row) is not claimed.
    """
    if not 1 <= limit <= _MAX_STRUCTURE_LIMIT:
        raise ValueError(f"limit must be between 1 and {_MAX_STRUCTURE_LIMIT}")
    c = classify_document(path)
    if c.kind != "spreadsheet" or c.protected or not c.readable:
        raise UnsupportedDocumentError("only a readable, unprotected XLSX structure can be read")
    if openpyxl is None:
        raise UnsupportedDocumentError("openpyxl is not installed")
    images_by_sheet, drawings_readable = _xlsx_drawing_report(path)
    try:
        workbook = openpyxl.load_workbook(path, data_only=False, read_only=False, keep_vba=c.macro_enabled)
    except (zipfile.BadZipFile, OSError, KeyError, ValueError) as exc:
        raise UnsupportedDocumentError(f"workbook could not be read safely: {type(exc).__name__}") from exc
    try:
        workbook_names = []
        for name, definition in workbook.defined_names.items():
            refers_to, was_clipped = _bounded_text(getattr(definition, "attr_text", None))
            workbook_names.append({"name": name, "scope": "workbook", "refers_to": refers_to,
                                   "refers_to_truncated": was_clipped,
                                   "kind": _defined_name_kind(refers_to or ""),
                                   "hidden": getattr(definition, "hidden", None)})
        sheet_names = []
        for ws in workbook.worksheets:
            for name, definition in ws.defined_names.items():
                refers_to, was_clipped = _bounded_text(getattr(definition, "attr_text", None))
                sheet_names.append({"name": name, "scope": ws.title, "refers_to": refers_to,
                                    "refers_to_truncated": was_clipped,
                                    "kind": _defined_name_kind(refers_to or ""),
                                    "hidden": getattr(definition, "hidden", None)})
        all_names = workbook_names + sheet_names
        titles = [ws.title for ws in workbook.worksheets] + \
                 [sheet.title for sheet in workbook.chartsheets]
        sheets = [_excel_sheet_structure(ws, (images_by_sheet or {}).get(ws.title, []), limit=limit)
                  for ws in list(workbook.worksheets)[:_MAX_STRUCTURE_SHEETS]]
        # Chart sheets hold no cells but still carry a chart inventory worth reporting.
        for chartsheet in list(workbook.chartsheets)[:_MAX_STRUCTURE_SHEETS - len(sheets)]:
            charts, chart_count = _chart_inventory(getattr(chartsheet, "_charts", []), limit=limit)
            sheets.append({"title": chartsheet.title, "state": getattr(chartsheet, "sheet_state", None),
                           "sheet_kind": "chartsheet", "charts": charts, "chart_count": chart_count,
                           "charts_truncated": chart_count > len(charts)})
        chart_count = sum(sheet.get("chart_count", 0) for sheet in sheets)
        image_count = sum(sheet.get("image_count", 0) for sheet in sheets)
        sheets_truncated = len(titles) > len(sheets)
        return {"classification": c.to_dict(), "readable": True, "sheet_titles": titles,
                "sheet_count": len(titles), "sheets": sheets,
                "sheets_truncated": sheets_truncated,
                "defined_names": all_names[:limit], "defined_name_count": len(all_names),
                "defined_names_truncated": len(all_names) > limit,
                "chart_count": chart_count, "image_count": image_count,
                # Totals only add up over the sheets that were detailed; an undetailed sheet
                # must not make a partial total look complete.
                "counts_cover_detailed_sheets_only": sheets_truncated,
                "drawings_readable": drawings_readable, "limit": limit,
                "note": "命名区域、条件格式、隐藏行列、合并区域与图表清单均按文件中的存储内容报告，"
                        "不求值也不应用规则；隐藏行列来自行高/列宽与 hidden 标记，"
                        "折叠分组或筛选造成的隐藏属于视图状态，这里不会声称检测到；"
                        "每一项都有对应总数与 *_truncated 标志，读不到绘图部件时 "
                        "drawings_readable 为 false，而不是报告“没有图片”；"
                        "工作表数超过上限时 sheets_truncated 为真，chart_count/image_count "
                        "只覆盖已详列的工作表。"}
    finally:
        workbook.close()


def _batch_xlsx_summary(path: Path, per_file: int) -> dict[str, Any]:
    """A light workbook summary; the handle is always closed, even on failure."""
    workbook = openpyxl.load_workbook(path, data_only=False, read_only=True)
    try:
        sheets, truncated = [], False
        for sheet in workbook.worksheets:
            if len(sheets) >= per_file:
                truncated = True
                break
            populated = sum(1 for row in sheet.iter_rows() for cell in row if cell.value is not None)
            sheets.append({"title": sheet.title, "state": sheet.sheet_state,
                           "max_row": sheet.max_row, "max_column": sheet.max_column,
                           "populated_cells": populated})
        return {"sheet_count": len(workbook.worksheets), "sheets": sheets,
                "sheets_truncated": truncated}
    finally:
        workbook.close()


def _batch_file_report(path: Path, *, operation: str, query: str | None,
                       per_file: int) -> dict[str, Any]:
    """One file's bounded contribution to a batch report."""
    classification = classify_document(path)
    entry: dict[str, Any] = {"file": path.name, "kind": classification.kind,
                             "size_bytes": classification.metadata.get("size")}
    if operation == "inspect":
        report = inspect_file(path)
        entry.update({"readable": report["readable"], "risks": report["risks"],
                      "macros": report["macros"], "protected": report["protected"]})
        return entry
    if operation == "summary":
        if classification.kind == "word":
            summary = summarize_word_document(path, chunk_size=50, max_chunks=2, preview_chars=60)
            entry.update({"paragraph_count": summary["paragraph_count"],
                          "table_count": summary["table_count"],
                          "character_count": summary["character_count"],
                          "headings": [item["text"][:_BATCH_TEXT]
                                       for item in summary["headings"][:per_file]],
                          "headings_truncated": summary["headings_truncated"]
                          or len(summary["headings"]) > per_file,
                          "preview": [chunk["preview"][:_BATCH_TEXT]
                                      for chunk in summary["chunks"][:1]]})
        else:
            entry.update(_batch_xlsx_summary(path, per_file))
        return entry
    hits = search_document(path, query or "", max_results=per_file)
    entry.update({"hits": [{"location": hit.get("location", hit.get("sheet")),
                            "match": str(hit.get("match", ""))[:_BATCH_TEXT],
                            "coordinate": hit.get("coordinate")}
                           for hit in hits],
                  "hit_count": len(hits), "hits_truncated": len(hits) >= per_file})
    return entry


def _batch_checked_report(candidate: Path, operation: str, query: str | None,
                          per_file: int) -> tuple[str, dict[str, Any]]:
    """One independent read; failures stay attached to their original candidate."""
    try:
        report = inspect_file(candidate)
        if not report.get("readable") or report.get("protected") or report.get("risks"):
            return "skipped", {"file": candidate.name,
                               "reason": "结构识别报风险，未解析：" +
                                         ", ".join(report.get("risks") or ["protected/unreadable"])}
        return "files", _batch_file_report(candidate, operation=operation, query=query,
                                           per_file=per_file)
    except Exception as exc:  # One bad file must not cancel the remaining reads.
        return "errors", {"file": candidate.name,
                          "error": f"{type(exc).__name__}: {str(exc)[:160]}"}


def batch_document_report(selected: str | os.PathLike[str], *, validate: Callable[[str], Any],
                          operation: str = "summary", query: str | None = None,
                          limit: int = _BATCH_MAX_FILES,
                          per_file: int = _BATCH_MAX_PER_FILE) -> dict[str, Any]:
    """Read-only report over the *authorized siblings* of one selected document.

    This is the only place that enumerates a directory, and it stays deliberately narrow:
    the scan is one level deep inside the selected file's own directory, every candidate is
    passed through ``validate`` — the same guard the single-file tools use, so ``..``,
    device paths, reserved names, trailing dots, symlinks and any escape from the authorized
    directory are refused — and a refused candidate is **reported as skipped instead of
    being followed or silently dropped**. Nothing here writes: no file, directory or output
    copy is created. Legacy formats and files the package inspector flags are skipped with
    their reason, never parsed.
    """
    if operation not in {"inspect", "summary", "search"}:
        raise ValueError("operation must be inspect, summary or search")
    if operation == "search" and not query:
        raise ValueError("search requires a non-empty query")
    if not 1 <= limit <= _BATCH_MAX_FILES:
        raise ValueError(f"limit must be between 1 and {_BATCH_MAX_FILES}")
    if not 1 <= per_file <= _BATCH_MAX_PER_FILE:
        raise ValueError(f"per_file must be between 1 and {_BATCH_MAX_PER_FILE}")
    directory = Path(selected).parent
    skipped: list[dict[str, str]] = []
    files: list[dict[str, Any]] = []
    errors: list[dict[str, str]] = []
    truncated = False
    try:
        entries = sorted(directory.iterdir(), key=lambda item: item.name.lower())
    except OSError as exc:
        return {"directory": str(directory), "operation": operation, "files": [], "file_count": 0,
                "skipped": [], "skipped_count": 0,
                "errors": [{"file": str(directory),
                            "error": f"{type(exc).__name__}: {str(exc)[:160]}"}],
                "error_count": 1, "truncated": False, "limit": limit,
                "per_file_limit": per_file, "wrote_nothing": True,
                "note": "授权目录无法枚举；未读取任何文件，也未写入任何内容"}
    candidates: list[Path] = []
    for entry in entries:
        if len(candidates) >= limit:
            # Probe only for another eligible sibling; never extend skipped lists
            # beyond the selected batch or flag unrelated directories as truncation.
            if (not entry.name.startswith((".", "~$"))
                    and entry.suffix.lower() in _BATCH_SUFFIXES
                    and not _BATCH_OUTPUT.search(entry.name)
                    and not entry.is_dir()):
                truncated = True
                break
            continue
        if entry.name.startswith(".") or entry.name.startswith("~$"):
            continue                                   # audit dir, hidden files, Office locks
        suffix = entry.suffix.lower()
        if _BATCH_OUTPUT.search(entry.name):
            skipped.append({"file": entry.name, "reason": "本应用生成的输出副本，不作为批量输入"})
            continue
        if suffix in _BATCH_LEGACY_SUFFIXES:
            skipped.append({"file": entry.name,
                            "reason": f"遗留格式 {suffix} 仅识别不解析（需 Office COM 通道）"})
            continue
        if suffix not in _BATCH_SUFFIXES:
            continue                                   # unrelated files are not "skipped inputs"
        try:
            candidates.append(Path(validate(str(entry))))
        except Exception as exc:                       # path guard refusal, recorded verbatim
            skipped.append({"file": entry.name,
                            "reason": f"路径护栏拒绝：{type(exc).__name__}: {str(exc)[:120]}"})
    if candidates:
        # Validate serially above, then read independent siblings with a fixed bound.
        # Executor.map returns in candidate order even when workers finish out of order.
        with ThreadPoolExecutor(max_workers=min(_BATCH_MAX_WORKERS, len(candidates))) as pool:
            outcomes = pool.map(lambda candidate: _batch_checked_report(candidate, operation,
                                                                         query, per_file), candidates)
            for category, item in outcomes:
                if category == "files":
                    files.append(item)
                elif category == "skipped":
                    skipped.append(item)
                else:
                    errors.append(item)
    return {"directory": str(directory), "operation": operation,
            "files": files, "file_count": len(files),
            "skipped": skipped, "skipped_count": len(skipped),
            "errors": errors, "error_count": len(errors),
            "truncated": truncated, "limit": limit, "per_file_limit": per_file,
            "wrote_nothing": True,
            "note": "只读批量：仅枚举所选文件所在目录的下一层，每个候选都经同一条路径护栏校验，"
                    "被拒绝的候选列在 skipped 而不是被跟随；遗留格式与报风险的文件不解析；"
                    "本操作不创建、不修改任何文件或目录；truncated 为真表示目录内还有未处理的候选。"}


def inspect_file(path: str | os.PathLike[str]) -> dict[str, Any]:
    """Compatibility-safe package inspection report used by the write guard.

    This function intentionally reports only structural safety signals; it does
    not parse or execute macros, links, embedded objects, or external targets.
    """
    c = classify_document(path)
    risks = list(c.suspicious_features)
    risks.extend(c.limitations)
    if c.kind == "legacy":
        risks.append("legacy_format_unsupported")
    return {
        "path": c.path,
        "kind": c.kind,
        "format": c.format,
        "readable": c.readable,
        "protected": c.protected if c.format == "ooxml" else None,
        "macros": c.macro_enabled if c.format == "ooxml" else None,
        "external_links": c.external_relationships if c.format == "ooxml" else None,
        "images": c.metadata.get("images"), "tables": c.metadata.get("tables"),
        "formulas": c.metadata.get("formulas"), "charts": c.metadata.get("charts"),
        "size_bytes": c.metadata.get("size"),
        "risks": sorted(set(risks)),
        "metadata": dict(c.metadata),
        "limitations": list(c.limitations),
    }


def inspect_document(path: str | os.PathLike[str], **kwargs: Any) -> dict[str, Any]:
    c = classify_document(path)
    if c.kind == "word": return inspect_docx(path, **kwargs)
    if c.kind == "spreadsheet": return inspect_xlsx(path)
    return {"classification": c.to_dict(), "readable": False, "limitations": list(c.limitations)}


def search_document(path: str | os.PathLike[str], query: str, **kwargs: Any) -> list[dict[str, Any]]:
    c = classify_document(path)
    if c.kind == "word":
        return search_docx(path, query, **kwargs)
    if c.kind != "spreadsheet" or openpyxl is None or not query:
        return []
    if c.protected or not c.readable:
        return []
    case_sensitive = kwargs.get("case_sensitive", False)
    max_results = kwargs.get("max_results", 50)
    if max_results <= 0:
        return []
    q = query if case_sensitive else query.lower()
    hits: list[dict[str, Any]] = []
    # Sparse sheets can report XFD1048576 as their used range. Iterating the
    # bounding rectangle would create billions of empty cells; inspect only
    # physically populated cells in the loaded workbook instead.
    formulas = openpyxl.load_workbook(path, data_only=False, read_only=False, keep_vba=c.macro_enabled)
    try:
        cached = openpyxl.load_workbook(path, data_only=True, read_only=False, keep_vba=c.macro_enabled)
        try:
            for ws in formulas.worksheets:
                cached_ws = cached[ws.title]
                for formula_cell in ws._cells.values():
                    cached_cell = cached_ws._cells.get((formula_cell.row, formula_cell.column))
                    formula_value = formula_cell.value
                    cached_value = cached_cell.value if cached_cell is not None else None
                    is_formula = formula_cell.data_type == "f"
                    candidates = []
                    if is_formula:
                        candidates.append(("formula_text", formula_value))
                        if cached_value is not None:
                            candidates.append(("cached_value", cached_value))
                    elif formula_value is not None:
                        candidates.append(("value", formula_value))
                    for matched_field, value in candidates:
                        text = str(value)
                        hay = text if case_sensitive else text.lower()
                        if q in hay:
                            hit = {"sheet": ws.title, "coordinate": formula_cell.coordinate,
                                   "text": text, "matched_field": matched_field}
                            if is_formula:
                                hit["formula_text"] = formula_value
                                hit["cached_value"] = cached_value
                            hits.append(hit)
                            break
                    if len(hits) >= max_results:
                        return hits
        finally:
            cached.close()
    finally:
        formulas.close()
    return hits


def file_fingerprint(path: str | os.PathLike[str]) -> str:
    """Return a stable content fingerprint for audit logs without changing the file."""
    digest = hashlib.sha256()
    with open(path, "rb") as source:
        for block in iter(lambda: source.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()
