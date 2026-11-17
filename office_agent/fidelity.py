from __future__ import annotations

from dataclasses import asdict
from contextlib import closing
import difflib
import hashlib
import json
import math
import os
from pathlib import Path
import stat
import sys
import tempfile
from xml.etree import ElementTree as ET
import zipfile

from .diagram_process import _run_owned
from .writes import Workspace, digest


MAX_FILE_BYTES = 64 * 1024 * 1024
MAX_PART_BYTES = 16 * 1024 * 1024
MAX_PACKAGE_BYTES = 128 * 1024 * 1024
MAX_PAGES = 40
MAX_PAGE_CHARS = 20000
MAX_PAGE_PIXELS = 8_000_000
RENDER_DPI = 144


class FidelityError(ValueError):
    def __init__(self, code):
        self.code = code
        super().__init__(code)


def bounded_bytes(path: Path) -> bytes:
    if path.is_symlink() or not path.is_file():
        raise FidelityError("ordinary_file_required")
    info = path.stat()
    if info.st_size > MAX_FILE_BYTES:
        raise FidelityError("file_size_limit")
    with path.open("rb") as stream:
        data = stream.read(MAX_FILE_BYTES + 1)
    if len(data) > MAX_FILE_BYTES:
        raise FidelityError("file_size_limit")
    return data


def package_projection(path: Path, *, for_render=False) -> dict[str, str]:
    if path.suffix.lower() not in {".docx", ".xlsx", ".pptx"}:
        raise FidelityError("unsupported_format")
    bounded_bytes(path)
    projection = {}
    try:
        with zipfile.ZipFile(path) as archive:
            parts = archive.infolist()
            names = [part.filename for part in parts]
            required = {'.docx': 'word/document.xml', '.xlsx': 'xl/workbook.xml', '.pptx': 'ppt/presentation.xml'}[path.suffix.lower()]
            if (not 1 <= len(parts) <= 2000 or len(names) != len(set(names))
                    or sum(part.file_size for part in parts) > MAX_PACKAGE_BYTES
                    or required not in names or "[Content_Types].xml" not in names):
                raise FidelityError("invalid_package")
            for part in parts:
                if (part.file_size > MAX_PART_BYTES or part.flag_bits & 1
                        or ".." in Path(part.filename).parts or part.filename.startswith(("/", "\\"))
                        or stat.S_ISLNK(part.external_attr >> 16)):
                    raise FidelityError("unsafe_package_part")
                data = archive.read(part)
                name = part.filename.lower()
                if name.endswith((".xml", ".rels")):
                    markup = data.replace(b"\x00", b"").lower()
                    if b"<!doctype" in markup or b"<!entity" in markup:
                        raise FidelityError("xml_entities_rejected")
                    root = ET.fromstring(data)
                    if for_render:
                        if any(node.attrib.get("TargetMode", "").lower() == "external"
                               for node in root.iter()):
                            raise FidelityError("external_relationship_rejected")
                        for node in root.iter():
                            tag = node.tag.rsplit("}", 1)[-1]
                            if tag in {"instrText", "fldSimple"}:
                                instruction = node.text if tag == 'instrText' else node.attrib.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}instr', '')
                                if (instruction or '').strip() != 'TOC \\o "1-3" \\h \\z \\u':
                                    raise FidelityError("active_content_rejected")
                            if tag in {"altChunk", "oleObject", "OLEObject"}:
                                raise FidelityError("active_content_rejected")
                            if tag == "f" and any(token in (node.text or "").upper()
                                                  for token in ("[", "]", "!", "|", "WEBSERVICE", "RTD", "DDE")):
                                raise FidelityError("external_formula_rejected")
                    data = ET.canonicalize(data).encode("utf-8")
                if for_render and any(token in name for token in (
                        "vbaproject", "activex", "embeddings/", "externallinks/", "connections", "customui")):
                    raise FidelityError("active_content_rejected")
                projection[part.filename] = hashlib.sha256(data).hexdigest()
        if for_render:
            from .rag_sources import _package, SourceError
            try:
                _package(bounded_bytes(path), path.suffix.lower())
            except SourceError:
                raise FidelityError('unsafe_relationship_or_package') from None
    except (OSError, zipfile.BadZipFile, ET.ParseError, RuntimeError, UnicodeError):
        raise FidelityError("invalid_package") from None
    return projection


def compare_structure(source: Path, output: Path) -> dict:
    if source.suffix.lower() != output.suffix.lower():
        raise FidelityError("format_mismatch")
    before, after = package_projection(source), package_projection(output)
    changed = sorted(name for name in before.keys() | after.keys() if before.get(name) != after.get(name))
    return {"status": "different" if changed else "equivalent", "equivalent": not changed,
            "scope": "all canonical XML and binary OOXML parts; excludes ZIP container metadata",
            "source_parts": len(before), "output_parts": len(after),
            "changed_parts": changed[:100], "changed_parts_count": len(changed),
            "truncated": len(changed) > 100}


def pdf_snapshot(path: Path, *, reading_order=False) -> dict:
    import pypdfium2 as pdfium

    data = bounded_bytes(path)
    if not data.startswith(b"%PDF-"):
        raise FidelityError("invalid_pdf")
    pages = []
    try:
        with pdfium.PdfDocument(data) as document:
            if not 1 <= len(document) <= MAX_PAGES:
                raise FidelityError("page_count_limit")
            for page in document:
                with closing(page):
                    width, height = page.get_size()
                    if (not all(math.isfinite(value) and value > 0 for value in (width, height))
                            or math.ceil(width * 2) * math.ceil(height * 2) > MAX_PAGE_PIXELS):
                        raise FidelityError("page_pixel_limit")
                    with closing(page.get_textpage()) as textpage:
                        count = textpage.count_chars()
                        if count > MAX_PAGE_CHARS:
                            raise FidelityError("page_text_limit")
                        text = textpage.get_text_range(0, count, errors="strict") if count else ""
                        if len(text) > MAX_PAGE_CHARS:
                            raise FidelityError("page_text_limit")
                        from .layout_checks import character_geometry
                        geometry = character_geometry(textpage, count, width, height)
                        from .pdf_reading_order import reading_text
                        ordered = reading_text(textpage,text,count,width,height) if reading_order and not geometry['unknown_boxes'] and not geometry['outside_page_characters'] else None
                    with closing(page.render(scale=RENDER_DPI / 72)) as bitmap:
                        with bitmap.to_pil() as picture:
                            with picture.convert("RGB") as rgb:
                                pixel_hash = hashlib.sha256(rgb.tobytes()).hexdigest()
                                pixels = list(rgb.size)
                    pages.append({"size_points": [width, height], "size_pixels": pixels,
                                  "text": text.replace("\r\n", "\n"), "pixel_sha256": pixel_hash, 'geometry': geometry,'reading_text':ordered})
    except FidelityError:
        raise
    except Exception:
        raise FidelityError("pdf_parse_failed") from None
    return {"page_count": len(pages), "pages": pages, "dpi": RENDER_DPI,
            "engine": "PDFium", "engine_version": str(pdfium.PDFIUM_INFO.version),
            "wrapper_version": str(pdfium.PYPDFIUM_INFO.version)}


def compare_pdf_snapshots(before: dict, after: dict) -> dict:
    from .layout_checks import compare_layout
    changes = []
    for index in range(max(before["page_count"], after["page_count"])):
        left = before["pages"][index] if index < before["page_count"] else None
        right = after["pages"][index] if index < after["page_count"] else None
        text_equal = left is not None and right is not None and left["text"] == right["text"]
        size_equal = left is not None and right is not None and left["size_points"] == right["size_points"]
        pixel_equal = size_equal and left["size_pixels"] == right["size_pixels"] and left["pixel_sha256"] == right["pixel_sha256"]
        diff = list(difflib.unified_diff((left or {}).get("text", "").splitlines(),
                                       (right or {}).get("text", "").splitlines(), n=1))
        excerpt = "\n".join(diff[:40])
        changes.append({"page": index + 1, "text_equal": text_equal, "size_equal": size_equal,
                        "pixels_equal": pixel_equal, "text_diff": excerpt[:2000],
                        "diff_truncated": len(diff) > 40 or len(excerpt) > 2000})
    equivalent = all(item["text_equal"] and item["pixels_equal"] for item in changes)
    return {"status": "equivalent" if equivalent else "different", "equivalent": equivalent,
            "scope": f"PDFium RGB pixels at {RENDER_DPI} DPI, page dimensions, page count and text",
            "source_pages": before["page_count"], "output_pages": after["page_count"],
            "page_count_equal": before["page_count"] == after["page_count"],
            "text_equal": all(item["text_equal"] for item in changes),
            "pixels_equal": all(item["pixels_equal"] for item in changes),
            "pages": changes, 'layout_observations': compare_layout(before, after), "engine": before["engine"], "engine_version": before["engine_version"],
            "wrapper_version": before["wrapper_version"], "dpi": RENDER_DPI,
            "limitations": ["Fixed renderer, fonts and DPI only; not Microsoft Office visual equivalence",
                            "Changes are observations, not approval or proof that an edit was intended"]}


def runtime_path(explicit: Path | None = None) -> Path:
    configured = explicit or os.environ.get("OFFICE_LIBREOFFICE_PATH")
    candidate = Path(configured) if configured else Path(__file__).resolve().parents[1] / ".runtime" / "libreoffice" / "program" / "soffice.com"
    if not candidate.is_absolute() or not candidate.is_file() or candidate.is_symlink():
        raise FidelityError("libreoffice_unavailable")
    return candidate.resolve()


def _temporary_root() -> str:
    if os.name == "nt":
        import win32api

        return win32api.GetShortPathName(str(Path(__file__).resolve().parents[1]))
    return tempfile.gettempdir()


def _profile(directory: Path) -> None:
    user = directory / "user"
    user.mkdir(parents=True)
    (user / "registrymodifications.xcu").write_text(
        '<?xml version="1.0" encoding="UTF-8"?>'
        '<oor:items xmlns:oor="http://openoffice.org/2001/registry">'
        '<item oor:path="/org.openoffice.Office.Common/Security/Scripting">'
        '<prop oor:name="MacroSecurityLevel" oor:op="fuse"><value>3</value></prop></item>'
        '<item oor:path="/org.openoffice.Office.Common/Save/Document">'
        '<prop oor:name="UpdateDocMode" oor:op="fuse"><value>0</value></prop></item>'
        '</oor:items>', encoding="utf-8")


def _run(command: tuple[str, ...], timeout_ms: int, runs: list) -> None:
    try:
        result = _run_owned(command, timeout_ms=timeout_ms)
    except Exception:
        raise FidelityError("worker_launch_or_cleanup_failed") from None
    runs.append(asdict(result))
    if result.code != "process_completed" or result.exit_code != 0 or result.empty_confirmed is not True:
        raise FidelityError("worker_timeout" if result.code == "render_timeout" else "worker_failed")


def _convert(source: Path, destination: Path, executable: Path, timeout_ms: int, runs: list) -> Path:
    destination.mkdir()
    profile = destination / "profile"
    _profile(profile)
    command = (sys.executable, str(Path(__file__).with_name("fidelity_worker.py")),
               "--convert", str(source), str(destination), str(executable))
    _run(command, timeout_ms, runs)
    result = destination / (source.stem + ".pdf")
    bounded_bytes(result)
    return result


def compare_documents(root: Path, source: Path, output: Path, *, executable: Path | None = None,
                      timeout_ms: int = 30000, static_docx: bool = False, range_anchors: list | None = None,
                      reuse_verified: bool = False) -> dict:
    if type(timeout_ms) is not int or not 100 <= timeout_ms <= 30000:
        raise FidelityError("invalid_timeout")
    space = Workspace(root)
    sources = [space.file(str(path)) for path in (source, output)]
    payloads = [bounded_bytes(path) for path in sources]
    hashes = [hashlib.sha256(data).hexdigest() for data in payloads]
    report = {"schema_version": 1, "source_sha256": hashes[0], "output_sha256": hashes[1],
              "inputs_unchanged": None, "structure": {"status": "unverified", "equivalent": None},
              "render": {"status": "unverified", "equivalent": None}, "processes": []}
    cache_payloads = []
    try:
        with tempfile.TemporaryDirectory(prefix="ofd-", dir=_temporary_root()) as raw:
            scratch = Path(raw)
            copies = []
            for index, (path, data) in enumerate(zip(sources, payloads)):
                copy = scratch / (f"input-{index}" + path.suffix.lower())
                copy.write_bytes(data)
                copies.append(copy)
            report["structure"] = compare_structure(*copies)
            if static_docx:
                from .range_edits import static_render_parts
                report['static_projection_sha256'] = []
                for path in copies:
                    package_projection(path)
                    with zipfile.ZipFile(path) as archive:
                        original = {name: archive.read(name) for name in archive.namelist()}
                    projected = static_render_parts(original)
                    report['static_projection_sha256'].append(hashlib.sha256(json.dumps(
                        {name: hashlib.sha256(data).hexdigest() for name, data in projected.items()}, sort_keys=True).encode('utf8')).hexdigest())
                    with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as archive:
                        for name, data in projected.items():
                            archive.writestr(name, data)
                report['render_projection'] = 'static field display and hyperlink text; original package and relationships preserved in output; dynamic field refresh not verified'
            runtime = runtime_path(executable)
            report["runtime_sha256"] = digest(runtime)
            if static_docx:
                from .render_identity import environment_identity
                from .render_handoff import implementation_identity
                report['conversion_identity'] = implementation_identity()
                report['render_environment'] = environment_identity(runtime)
                report['anchor_contexts'] = range_anchors
            for path in copies:
                package_projection(path, for_render=True)
            pdfs = []
            cache_keys = []
            if reuse_verified:
                report['verified_render_reuse'] = []
            for index, path in enumerate(copies):
                identifier = None
                cached = None
                if reuse_verified and static_docx and report['render_environment']['complete']:
                    from . import verified_render_cache
                    identifier = verified_render_cache.key(space.root, hashes[index], report['static_projection_sha256'][index],
                        report['runtime_sha256'], report['render_environment'], report['conversion_identity'])
                    cached = verified_render_cache.get(identifier)
                cache_keys.append(identifier)
                if cached is not None:
                    pdf = scratch / f'verified-{index}.pdf'
                    pdf.write_bytes(cached)
                else:
                    pdf = _convert(path, scratch / f"pdf-{index}", runtime, timeout_ms, report["processes"])
                pdfs.append(pdf)
                if reuse_verified:
                    report['verified_render_reuse'].append({'input_sha256': hashes[index], 'reused': cached is not None})
            result_file = scratch / "render.json"
            command = (sys.executable, str(Path(__file__).with_name("fidelity_worker.py")),
                       str(pdfs[0]), str(pdfs[1]), str(result_file))
            if range_anchors is not None:
                encoded = json.dumps(range_anchors, ensure_ascii=False).encode('utf-8')
                if len(encoded) > 60000 or len(range_anchors) > 12:
                    raise FidelityError('anchor_capacity')
                anchors_path = scratch / 'anchors.json'
                anchors_path.write_bytes(encoded)
                command += (str(anchors_path),)
            _run(command, timeout_ms, report["processes"])
            result = bounded_bytes(result_file)
            if len(result) > 200000:
                raise FidelityError("report_size_limit")
            report["render"] = json.loads(result)
            if static_docx and report['render_environment']['complete'] and environment_identity(runtime) != report['render_environment']:
                raise FidelityError('render_environment_changed')
            if reuse_verified and report['render'].get('status') in {'equivalent', 'different'}:
                cache_payloads = [(identifier, pdf.read_bytes()) for identifier, pdf in zip(cache_keys, pdfs) if identifier]
            if static_docx and range_anchors and all(space.file(str(path)) == path and digest(path) == expected
                                                     for path, expected in zip(sources, hashes)):
                from .render_handoff import offer
                report['inputs_unchanged'] = True
                offer(space.root, report, pdfs)
    except FidelityError as error:
        report["render"] = {"status": "unavailable" if error.code == "libreoffice_unavailable" else "error",
                            "equivalent": None, "code": error.code}
    finally:
        report["inputs_unchanged"] = all(space.file(str(path)) == path and digest(path) == expected
                                         for path, expected in zip(sources, hashes))
        if not report["inputs_unchanged"]:
            from .render_handoff import clear
            clear()
            report["render"] = {"status": "error", "equivalent": None, "code": "inputs_changed"}
            report["structure"] = {"status": "error", "equivalent": None, "code": "inputs_changed"}
            if reuse_verified:
                from .verified_render_cache import clear
                clear()
    if reuse_verified and report['inputs_unchanged'] and report['render'].get('status') in {'equivalent', 'different'}:
        from .verified_render_cache import offer
        for identifier, data in cache_payloads:
            offer(identifier, data)
    return report
