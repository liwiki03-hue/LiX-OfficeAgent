"""Deterministic, allowlisted, preview-first OOXML edits. Never evaluates model code."""
from __future__ import annotations

import re
import zipfile
from copy import deepcopy
from io import BytesIO
from xml.etree import ElementTree as ET
import hashlib
import json
import math
import os
import shutil
import stat
import time
import uuid
from datetime import datetime, timezone
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Literal

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.table import _Cell
from openpyxl import load_workbook
from openpyxl.chart import BarChart, LineChart, PieChart, Reference
from openpyxl.formula.tokenizer import Tokenizer, TokenizerError
from openpyxl.formula.translate import Translator, TranslatorError
from openpyxl.utils.cell import get_column_letter, quote_sheetname, range_boundaries
from pydantic import BaseModel, ConfigDict, Field, model_validator


class GuardError(ValueError):
    pass


class Strict(BaseModel):
    model_config = ConfigDict(extra="forbid", strict=True)


class WordReplace(Strict):
    # ``kind`` labels the request in previews and audit records. It is deliberately a plain
    # string rather than a validation gate: the calling tool name selects the request type
    # (see app._PLAN_MODELS), so a model echoing the wrong label must not turn an otherwise
    # valid call into a rejection. All semantic fields below stay strictly validated.
    kind: str = "word_replace"
    path: str
    old: str = Field(min_length=1, max_length=500)
    new: str = Field(max_length=500)
    paragraph_index: int = Field(ge=0)
    expected_count: int = Field(default=1, ge=1, le=100)


class WordBatchReplace(Strict):
    """Replace a literal everywhere in the body, capped and requiring its own confirmation."""
    kind: str = "word_batch_replace"
    path: str
    old: str = Field(min_length=1, max_length=200)
    new: str = Field(max_length=200)
    max_occurrences: int = Field(default=20, ge=1, le=100)


class WordFormat(Strict):
    kind: str = "word_format"
    path: str
    paragraph_index: int = Field(ge=0)
    run_index: int = Field(ge=0)
    bold: bool | None = None
    italic: bool | None = None
    underline: bool | None = None

    @model_validator(mode="after")
    def check_format(self):
        if all(value is None for value in (self.bold, self.italic, self.underline)):
            raise ValueError("at least one formatting property is required")
        return self


class WordTableCellReplace(Strict):
    kind: str = "word_table_cell_replace"
    path: str
    table_index: int = Field(ge=0)
    row_index: int = Field(ge=0)
    cell_index: int = Field(ge=0)
    paragraph_index: int = Field(ge=0)
    old: str = Field(min_length=1, max_length=500)
    new: str = Field(max_length=500)


class WordMergedCellText(Strict):
    """Replace or append text in the cell that carries it, inside a merged table too.

    ``row_index``/``column_index`` are logical grid coordinates, not the physical ``w:tc``
    position. A coordinate covered by a gridSpan or continued by a vMerge is refused with
    the origin coordinate, because text written there is not the text Word displays.
    """
    kind: str = "word_merged_cell_text"
    path: str
    table_index: int = Field(ge=0)
    row_index: int = Field(ge=0)
    column_index: int = Field(ge=0)
    action: Literal["replace", "append"]
    old: str = Field(default="", max_length=500)
    new: str = Field(min_length=1, max_length=500)
    paragraph_index: int = Field(default=0, ge=0)
    expected_count: int = Field(default=1, ge=1, le=100)

    @model_validator(mode="after")
    def check_action(self):
        if self.action == "replace":
            if not self.old:
                raise ValueError("replace needs old: the exact literal the cell holds now")
            if self.old == self.new:
                raise ValueError("replacement would make no change")
        elif self.old:
            raise ValueError("append adds one paragraph at the end and takes no old literal")
        return self


class WordRevisionDecision(Strict):
    """Accept or reject every content revision, body only (word/document.xml).

    Accepting unwraps ``w:ins`` (its text becomes ordinary text) and removes ``w:del``
    together with the deleted content. Rejecting removes ``w:ins`` and unwraps ``w:del``,
    turning ``w:delText`` back into ``w:t``. Formatting revisions (``w:rPrChange`` /
    ``w:pPrChange``) are settled with the same decision; every other revision mark —
    moves, row/table/cell revisions, paragraph-mark revisions — refuses the whole edit
    with a precise list, so nothing is ever left half-decided.
    """
    kind: str = "word_revision_decision"
    path: str
    action: Literal["accept_all", "reject_all"]


class DocxMerge(Strict):
    """Merge several DOCX files from the anchor's own directory into one **new** output.

    ``path`` is the file the user authorized; ``sources`` are *file names* resolved inside
    that file's directory, so a request can never name a path elsewhere. No source is
    modified: only a new output file is created, next to the anchor.
    """
    kind: str = "docx_merge"
    path: str
    sources: list[str] = Field(min_length=2, max_length=20)

    @model_validator(mode="after")
    def check_sources(self):
        if len(set(self.sources)) != len(self.sources):
            raise ValueError("merge sources must be listed once each, in order")
        for name in self.sources:
            if not isinstance(name, str) or not name or len(name) > 200:
                raise ValueError("each merge source must be a file name")
            if name != Path(name).name or name in {".", ".."} or "\\" in name or "/" in name:
                raise ValueError(f"merge sources must be file names in the anchor's directory, "
                                 f"not paths: {name!r}")
            if not name.lower().endswith(".docx"):
                raise ValueError(f"only .docx files can be merged: {name!r}")
        return self


class DocxSplit(Strict):
    """Split one DOCX into several **new** DOCX files, one per top-level section heading.

    ``level`` is the deepest heading level that starts a new part. The source is never
    modified; every output is a new file next to it.
    """
    kind: str = "docx_split"
    path: str
    level: int = Field(default=1, ge=1, le=3)


class LegacyWordReplace(Strict):
    """Replace one literal inside one body paragraph of a legacy ``.doc`` via Office COM.

    The legacy format cannot be edited by python-docx, so the edit runs through an isolated
    Office instance — and only when the capability probe for that instance has succeeded.
    """
    kind: str = "legacy_word_replace"
    path: str
    paragraph_index: int = Field(default=0, ge=0)
    old: str = Field(min_length=1, max_length=500)
    new: str = Field(max_length=500)
    expected_count: int = Field(default=1, ge=1, le=100)

    @model_validator(mode="after")
    def check_change(self):
        if self.old == self.new:
            raise ValueError("replacement would make no change")
        return self


class LegacyCellSet(Strict):
    """Assign one literal value to one ``.xls`` cell via Office COM.

    ``expected_original`` states the value the caller believes is there now (``None`` for a
    blank cell); the edit is refused unless it matches exactly.
    """
    kind: str = "legacy_cell_set"
    path: str
    sheet: str = Field(min_length=1)
    cell: str
    value: str | int | float | bool | None
    expected_original: str | int | float | bool | None

    @model_validator(mode="after")
    def check_cell(self):
        _check_excel_cell_address(self.cell)
        if isinstance(self.value, str) and self.value.startswith("="):
            raise ValueError("legacy cell writes accept literal values only, never formulas")
        for candidate in (self.value, self.expected_original):
            if isinstance(candidate, float) and not math.isfinite(candidate):
                raise ValueError("non-finite numbers are not valid Excel literal values")
        return self


_MAX_LEGACY_PARAGRAPH_WINDOW = 100   # Paragraphs read/verified per legacy .doc operation.
_LEGACY_REGION = "A1:T20"            # Bounded window verified around one legacy .xls cell.


def _legacy_same_value(current: Any, expected: Any) -> bool:
    """Compare an Office value with a declared one, normalising whole numbers.

    Office returns every number as a float, so 42 and 42.0 must compare equal while booleans
    stay distinct from 1/0 and text stays exact.
    """
    if isinstance(current, bool) or isinstance(expected, bool):
        return isinstance(current, bool) and isinstance(expected, bool) and current is expected
    if current is None or expected is None:
        return current is None and expected is None
    if isinstance(current, (int, float)) and isinstance(expected, (int, float)):
        return float(current) == float(expected)
    return type(current) is type(expected) and current == expected


def _require_legacy_source(src: Path) -> None:
    """Only a genuine OLE compound file may enter the COM write channel."""
    from .documents import OLE_MAGIC, classify_document
    classification = classify_document(src)
    if classification.kind != "legacy" or classification.format != "ole_compound":
        raise GuardError("legacy edit requires a real OLE .doc/.xls file")
    with src.open("rb") as stream:
        if stream.read(8) != OLE_MAGIC:
            raise GuardError("legacy edit requires the OLE compound-file header")


def _legacy_paragraph_window(backend: Any, path: Path) -> list[str]:
    paragraphs = backend.read_word_paragraphs(str(path), limit=_MAX_LEGACY_PARAGRAPH_WINDOW)
    if not isinstance(paragraphs, list) or any(not isinstance(item, str) for item in paragraphs):
        raise GuardError("Office backend returned an unusable paragraph list")
    return paragraphs


def _legacy_region_values(backend: Any, path: Path, sheet: str) -> dict[str, Any]:
    first, last = _LEGACY_REGION.split(":")
    values = backend.read_excel_region(str(path), sheet=sheet, first_cell=first, last_cell=last)
    if not isinstance(values, dict):
        raise GuardError("Office backend returned an unusable region")
    return values


def _legacy_verify_window(before: dict[str, Any], after: dict[str, Any], *,
                          excluded: str, label: str) -> None:
    """Everything in the verified window except the edited cell must be unchanged."""
    if set(before) != set(after):
        raise GuardError(f"legacy 非目标{label}集合发生变化")
    changed = [key for key in before
               if key != excluded.upper() and not _legacy_same_value(before[key], after[key])]
    if changed:
        raise GuardError(f"legacy 非目标{label}发生变化：{', '.join(sorted(changed)[:5])}")


_MERGE_MAX_SOURCES = 20      # One confirmed merge handles at most this many documents.
_MERGE_MAX_PARTS = 20        # One confirmed split creates at most this many outputs.
_MERGE_PREVIEW_TEXTS = 10    # Non-empty paragraph texts shown in a merge/split preview.


def _merge_sources(space: Workspace, anchor: Path, names: list[str]) -> list[Path]:
    """Resolve merge source *names* inside the anchor's directory, guarded one by one.

    Names are deliberately not paths: a request cannot walk out of the directory, and every
    resolved candidate still passes through the same ``Workspace.file`` guard as a
    single-file tool, so symlinks, reserved names and trailing dots are refused there too.
    """
    directory = anchor.parent
    resolved: list[Path] = []
    for name in names:
        candidate = directory / name
        try:
            resolved.append(space.file(str(candidate)))
        except GuardError:
            raise
        except Exception as exc:
            raise GuardError(f"merge source {name!r} was refused: {type(exc).__name__}") from exc
    return resolved


def _b7_numbered_count(document: Any, element: Any) -> int:
    """Count effective numbered paragraphs, including numbering inherited from styles."""
    styles = {s.get(qn("w:styleId")): s for s in document.styles.element.findall(qn("w:style"))}
    default = next((sid for sid, s in styles.items() if s.get(qn("w:type")) == "paragraph"
                    and s.get(qn("w:default")) in {"1", "true", "on"}), None)
    count = 0
    for paragraph in element.iter(qn("w:p")):
        num = paragraph.find("./" + qn("w:pPr") + "/" + qn("w:numPr") + "/" + qn("w:numId"))
        ref = paragraph.find("./" + qn("w:pPr") + "/" + qn("w:pStyle"))
        sid = ref.get(qn("w:val")) if ref is not None else default
        seen = set()
        while num is None and sid in styles and sid not in seen:
            seen.add(sid)
            style = styles[sid]
            num = style.find("./" + qn("w:pPr") + "/" + qn("w:numPr") + "/" + qn("w:numId"))
            based = style.find(qn("w:basedOn"))
            sid = based.get(qn("w:val")) if based is not None else None
        if num is not None and num.get(qn("w:val")) != "0":
            count += 1
    return count


def _docx_signature(path: Path) -> dict[str, Any]:
    """A comparable description of one DOCX: non-empty texts, tables, sections and features.

    Empty paragraphs are excluded because a merge inserts empty section-break paragraphs
    between sources; everything that carries content is compared exactly. Relationship
    references and numbering are counted because they are the formatting that a copy can
    lose, and the preview has to say so rather than pretend the copy is complete.
    """
    document = Document(path)
    texts = [paragraph.text for paragraph in document.paragraphs if paragraph.text]
    body = document.element.body
    relationships = sum(1 for node in body.iter() for attribute in node.attrib
                        if attribute.startswith(_RELATIONSHIP_ATTR))
    return {"texts": texts, "tables": len(document.tables),
            "sections": len(document.sections), "relationships": relationships,
            "numbered": _b7_numbered_count(document, body)}


def _signature_value(texts: list[str], tables: int, sections: int) -> dict[str, Any]:
    return {"text_sha256": hashlib.sha256(
                json.dumps(texts, ensure_ascii=False).encode("utf-8")).hexdigest(),
            "text_count": len(texts), "tables": tables, "sections": sections,
            "preview": texts[:_MERGE_PREVIEW_TEXTS]}


def _merge_plan_value(sources: list[Path]) -> tuple[str, str, dict[str, int]]:
    """(inputs value, expected output value, feature totals) computed without writing anything."""
    preview_document = Document()
    from .documents import inspect_file
    for index, source in enumerate(sources, 1):
        report = inspect_file(source)
        if (report.get("risks") or report.get("protected") or report.get("macros")
                or report.get("external_links")):
            raise GuardError(f"B7 merge source requires a high-fidelity adapter: {source.name}; "
                             f"{report.get('risks', [])}")
        _migrate_b7(Document(source), preview_document, index)
    signatures = [_docx_signature(source) for source in sources]
    texts = [text for signature in signatures for text in signature["texts"]]
    tables = sum(signature["tables"] for signature in signatures)
    sections = sum(signature["sections"] for signature in signatures)
    built = {"fidelity_policy": dict(_B7_POLICY), "inputs": [{"name": source.name, "sha256": digest(source),
                         "text_count": len(signature["texts"]), "tables": signature["tables"],
                         "sections": signature["sections"],
                         "relationships": signature["relationships"],
                         "numbered": signature["numbered"]}
                        for source, signature in zip(sources, signatures)]}
    totals = {"relationships": sum(signature["relationships"] for signature in signatures),
              "numbered": sum(signature["numbered"] for signature in signatures)}
    expected = _signature_value(texts, tables, sections)
    return (json.dumps(built, ensure_ascii=False), json.dumps(expected, ensure_ascii=False), totals)


_RELATIONSHIP_ATTR = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"

_B7_POLICY = {"styles": "source-isolated IDs and names; references rewritten",
              "numbering": "source-isolated num/abstractNum IDs; references rewritten",
              "header_footer_footnotes_endnotes_comments": "refuse if present; accepted inputs absent",
              "verification": "OOXML structural only; visual fidelity not verified"}
_STYLE_REFS = {qn("w:" + name) for name in
               ("pStyle", "rStyle", "tblStyle", "basedOn", "next", "link",
                "numStyleLink", "styleLink")}


def _b7_guard(document: Any) -> None:
    """Fail closed before preview/output; never silently discard auxiliary content."""
    refused = ("header", "footer", "footnotes", "endnotes", "comments")
    for part in document.part.package.parts:
        name = str(part.partname).lower()
        if any(name.startswith("/word/" + kind) for kind in refused):
            raise GuardError(f"B7 policy refuses headers/footers/footnotes/endnotes/comments: {name}")
    forbidden = {qn("w:" + name) for name in
                 ("headerReference", "footerReference", "footnoteReference", "endnoteReference",
                  "commentReference", "commentRangeStart", "commentRangeEnd")}
    if any(node.tag in forbidden for node in document.element.iter()):
        raise GuardError("B7 policy refuses auxiliary-part references, including dangling references")
    for node in document.element.body.iter():
        for attr, value in node.attrib.items():
            if not attr.startswith(_RELATIONSHIP_ATTR):
                continue
            rel = document.part.rels.get(value)
            if rel is None or (not rel.is_external and not rel.reltype.endswith("/image")):
                raise GuardError(f"B7 refuses unresolved/unsupported body relationship: {value}")
    numbering = document.part.numbering_part.element
    if numbering.find(qn("w:numPicBullet")) is not None:
        raise GuardError("B7 refuses picture-bullet numbering")
    for root in (document.styles.element, numbering):
        if any(attr.startswith(_RELATIONSHIP_ATTR) for node in root.iter() for attr in node.attrib):
            raise GuardError("B7 refuses relationships in styles/numbering definitions")


def _xml_value(element: Any) -> bytes:
    return ET.tostring(element) if element is not None else b""


def _b7_environment(document: Any, target: Any) -> None:
    # Defaults and themes affect even explicitly styled text. Refuse differing global
    # contexts rather than claim that copying style definitions alone preserves them.
    if _xml_value(document.styles.element.find(qn("w:docDefaults"))) != _xml_value(
            target.styles.element.find(qn("w:docDefaults"))):
        raise GuardError("B7 refuses incompatible docDefaults; explicit style migration is insufficient")
    def themes(doc):
        return [rel.target_part.blob for rel in doc.part.rels.values()
                if rel.reltype.endswith("/theme") and not rel.is_external]
    if themes(document) != themes(target):
        raise GuardError("B7 refuses incompatible document themes")


def _rewrite_b7(element: Any, styles: dict[str, str], nums: dict[str, str],
                abstracts: dict[str, str]) -> None:
    for node in element.iter():
        mapping = (styles if node.tag in _STYLE_REFS else nums if node.tag == qn("w:numId")
                   else abstracts if node.tag == qn("w:abstractNumId") else None)
        if mapping is None:
            continue
        value = node.get(qn("w:val"))
        if node.tag == qn("w:numId") and value == "0":
            continue  # Word's explicit no-numbering sentinel, not a list definition.
        if value not in mapping:
            raise GuardError(f"B7 unresolved {node.tag.rsplit('}', 1)[-1]} reference: {value}")
        node.set(qn("w:val"), mapping[value])


def _migrate_b7(document: Any, target: Any, index: int) -> dict[str, Any]:
    """Namespace the entire source style/numbering graph, including inherited lists."""
    _b7_guard(document)
    _b7_environment(document, target)
    source_styles = list(document.styles.element.findall(qn("w:style")))
    styles = {style.get(qn("w:styleId")): f"B7s{index}_{n}"
              for n, style in enumerate(source_styles, 1)}
    if None in styles or len(styles) != len(source_styles):
        raise GuardError("B7 missing/duplicate style IDs")
    numbering = document.part.numbering_part.element
    target_numbering = target.part.numbering_part.element
    abstracts_nodes = list(numbering.findall(qn("w:abstractNum")))
    num_nodes = list(numbering.findall(qn("w:num")))
    def allocated(nodes, attr, existing):
        start = max([int(node.get(qn(attr))) for node in existing] + [0]) + 1
        result = {node.get(qn(attr)): str(start + n) for n, node in enumerate(nodes)}
        if None in result or len(result) != len(nodes):
            raise GuardError(f"B7 missing/duplicate {attr} IDs")
        return result
    abstracts = allocated(abstracts_nodes, "w:abstractNumId",
                          target_numbering.findall(qn("w:abstractNum")))
    nums = allocated(num_nodes, "w:numId", target_numbering.findall(qn("w:num")))
    defaults = {}
    expected_styles, expected_numbering = [], []
    for style in source_styles:
        copied = deepcopy(style)
        old_id = copied.get(qn("w:styleId"))
        copied.set(qn("w:styleId"), styles[old_id])
        if copied.get(qn("w:default")) in {"1", "true", "on"}:
            defaults[copied.get(qn("w:type"))] = styles[old_id]
        copied.attrib.pop(qn("w:default"), None)
        name = copied.find(qn("w:name"))
        if name is not None:
            name.set(qn("w:val"), f"{name.get(qn('w:val'))} [B7 source {index}; {styles[old_id]}]")
        aliases = copied.find(qn("w:aliases"))
        if aliases is not None:
            aliases.set(qn("w:val"), ",".join(
                f"{alias.strip()} [B7 {styles[old_id]}]"
                for alias in (aliases.get(qn("w:val")) or "").split(",") if alias.strip()))
        _rewrite_b7(copied, styles, nums, abstracts)
        target.styles.element.append(copied)
        expected_styles.append(_xml_value(copied))
    for nodes, attr, mapping in ((abstracts_nodes, "w:abstractNumId", abstracts),
                                 (num_nodes, "w:numId", nums)):
        for node in nodes:
            copied = deepcopy(node)
            copied.set(qn(attr), mapping[node.get(qn(attr))])
            if attr == "w:abstractNumId":
                nsid = copied.find(qn("w:nsid"))
                if nsid is not None:
                    nsid.set(qn("w:val"), hashlib.sha256(
                        f"B7:{index}:{node.get(qn(attr))}".encode()).hexdigest()[:8].upper())
            _rewrite_b7(copied, styles, nums, abstracts)
            # abstractNum must precede num in schema order.
            if attr == "w:abstractNumId":
                first_num = target_numbering.find(qn("w:num"))
                if first_num is not None:
                    first_num.addprevious(copied)
                else:
                    target_numbering.append(copied)
            else:
                target_numbering.append(copied)
            expected_numbering.append(_xml_value(copied))
    # Validate body refs during preview too, without mutating the source.
    _rewrite_b7(deepcopy(document.element.body), styles, nums, abstracts)
    return {"styles": styles, "nums": nums, "abstracts": abstracts, "defaults": defaults,
            "expected_styles": expected_styles, "expected_numbering": expected_numbering}


def _apply_b7(element: Any, migration: dict[str, Any]) -> None:
    _rewrite_b7(element, migration["styles"], migration["nums"], migration["abstracts"])
    for node in element.iter():
        if node.tag == qn("w:p"):
            props, style, kind = "w:pPr", "w:pStyle", "paragraph"
        elif node.tag == qn("w:tbl"):
            props, style, kind = "w:tblPr", "w:tblStyle", "table"
        elif node.tag == qn("w:r"):
            props, style, kind = "w:rPr", "w:rStyle", "character"
        else:
            continue
        default = migration["defaults"].get(kind)
        if default is None:
            continue
        properties = node.find(qn(props))
        if properties is None:
            properties = OxmlElement(props)
            node.insert(0, properties)
        if properties.find(qn(style)) is None:
            reference = OxmlElement(style)
            reference.set(qn("w:val"), default)
            properties.insert(0, reference)


def _verify_b7(output: Path, migrations: list[dict[str, Any]], expected_body: bytes,
               expected_numbered: int) -> None:
    reopened = Document(output)
    if _b7_numbered_count(reopened, reopened.element.body) != expected_numbered:
        raise GuardError("B7 reopened effective numbering count differs from copied paragraphs")
    if _xml_value(reopened.element.body) != expected_body:
        raise GuardError("B7 reopened body differs from copied formatting/reference graph")
    styles = {_xml_value(node) for node in reopened.styles.element.findall(qn("w:style"))}
    numbering = {_xml_value(node) for node in reopened.part.numbering_part.element}
    for migration in migrations:
        if not set(migration["expected_styles"]) <= styles or not set(
                migration["expected_numbering"]) <= numbering:
            raise GuardError("B7 reopened style/numbering graph differs from migrated definitions")
    _rewrite_b7(deepcopy(reopened.element.body),
                {n.get(qn("w:styleId")): n.get(qn("w:styleId"))
                 for n in reopened.styles.element.findall(qn("w:style"))},
                {n.get(qn("w:numId")): n.get(qn("w:numId"))
                 for n in reopened.part.numbering_part.element.findall(qn("w:num"))}, {})


def _remap_relationship(source_document: Any, target_document: Any, relation_id: str,
                        cache: dict[Any, str]) -> str | None:
    """Re-create one relationship of a copied block inside the target document.

    Copying XML without its relationships would leave ``r:embed``/``r:id`` pointing at ids
    the new package never defines, so images and hyperlinks would silently vanish. Returns
    the new id, or ``None`` when the relationship cannot be resolved (the caller refuses).
    """
    key = (source_document.part, relation_id)
    if key in cache:
        return cache[key]
    relationship = source_document.part.rels.get(relation_id)
    if relationship is None:
        return None
    if relationship.is_external:
        new_id = target_document.part.relate_to(relationship.target_ref,
                                               relationship.reltype, is_external=True)
    elif relationship.reltype.endswith("/image"):
        # Import by bytes: foreign packages may both call different images image1.png.
        new_id, _ = target_document.part.get_or_add_image(BytesIO(relationship.target_part.blob))
    else:
        raise GuardError(f"B7 refuses unsupported internal relationship: {relationship.reltype}")
    cache[key] = new_id
    return new_id


def _copy_block(source_document: Any, target_document: Any, element: Any,
                cache: dict[Any, str], report: dict[str, Any]) -> Any:
    """Deep-copy one block, remapping every relationship it references.

    Style and numbering references are migrated separately by B7; this function
    recreates image/hyperlink relationships and counts effective numbered paragraphs.
    """
    copied = deepcopy(element)
    for node in copied.iter():
        for attribute, value in list(node.attrib.items()):
            if not attribute.startswith(_RELATIONSHIP_ATTR):
                continue
            new_id = _remap_relationship(source_document, target_document, value, cache)
            if new_id is None:
                raise GuardError(f"源文档的关系 {value} 无法重建，拒绝复制以免静默丢失图片或链接")
            node.set(attribute, new_id)
    report["numbered_paragraphs"] += _b7_numbered_count(source_document, copied)
    return copied


def _merge_docx_documents(sources: list[Path], output: Path) -> dict[str, Any]:
    """Concatenate the block content of each DOCX into a new one, keeping its section setup.

    Each source keeps its own section settings: earlier sources end with a paragraph-level
    ``sectPr`` carrying their page setup, and the last source's body ``sectPr`` becomes the
    merged document's final section. Relationships are re-created for every copied block, so
    images and hyperlinks survive; sources are opened read-only and never saved.
    """
    merged = Document()
    body = merged.element.body
    report = {"relationships_remapped": 0, "numbered_paragraphs": 0, "sources": len(sources)}
    cache: dict[Any, str] = {}
    migrations = []
    for block in list(body):
        if block.tag != qn("w:sectPr"):
            body.remove(block)                     # drop the template's empty paragraph
    last_section = None
    for index, source in enumerate(sources):
        document = Document(source)
        migration = _migrate_b7(document, merged, index + 1)
        migrations.append(migration)
        source_body = document.element.body
        section = source_body.find(qn("w:sectPr"))
        last_section = section
        for child in source_body:
            if child is section:
                continue
            copied = _copy_block(document, merged, child, cache, report)
            _apply_b7(copied, migration)
            body.append(copied)
        if section is not None and index < len(sources) - 1:
            paragraph = OxmlElement("w:p")
            properties = OxmlElement("w:pPr")
            properties.append(deepcopy(section))
            paragraph.append(properties)
            body.append(paragraph)
    existing = body.find(qn("w:sectPr"))
    if existing is not None:
        body.remove(existing)
    if last_section is not None:
        body.append(deepcopy(last_section))
    merged.save(output)
    _verify_b7(output, migrations, _xml_value(body), report["numbered_paragraphs"])
    report.update({"styles_migrated": sum(len(m["styles"]) for m in migrations),
                   "numbering_definitions_migrated": sum(len(m["nums"]) for m in migrations),
                   "fidelity_policy": dict(_B7_POLICY)})
    report["relationships_remapped"] = len(cache)
    return report


def _split_parts(source: Path, level: int) -> list[tuple[str, list[Any]]]:
    """Split one DOCX body into parts at headings of ``level`` or shallower.

    Returns ``(title, block elements)`` in document order. Content before the first heading
    becomes its own part only when it carries something other than empty paragraphs.
    """
    document = Document(source)
    parts: list[tuple[str, list[Any]]] = []
    current_title: str | None = None
    current: list[Any] = []
    for block in document.iter_inner_content():
        style = getattr(getattr(block, "style", None), "name", "") or ""
        is_heading = style.lower().startswith("heading")
        depth = 0
        if is_heading:
            match = re.search(r"(\d+)", style)
            depth = int(match.group(1)) if match else 1
        if is_heading and depth <= level:
            if current and (current_title is not None or any(
                    getattr(item, "text", "") for item in current)):
                parts.append((current_title or "（标题之前的内容）", current))
            # The heading stays with its own part, so every output is self-describing.
            current_title, current = block.text, [block]
            continue
        current.append(block)
    if current and (current_title is not None or any(getattr(item, "text", "") for item in current)):
        parts.append((current_title or "（标题之前的内容）", current))
    return parts


def _write_split_part(blocks: list[Any], section: Any, output: Path,
                      source_document: Any) -> dict[str, Any]:
    """Write one part as a new DOCX, inheriting the source's page setup and relationships."""
    document = Document()
    body = document.element.body
    report = {"relationships_remapped": 0, "numbered_paragraphs": 0,
              "fidelity_policy": dict(_B7_POLICY)}
    cache: dict[Any, str] = {}
    migration = _migrate_b7(source_document, document, 1)
    for block in list(body):
        if block.tag != qn("w:sectPr"):
            body.remove(block)
    for block in blocks:
        element = getattr(block, "_p", None)
        if element is None:
            element = getattr(block, "_tbl", None)
        if element is not None:
            copied = _copy_block(source_document, document, element, cache, report)
            _apply_b7(copied, migration)
            body.append(copied)
    existing = body.find(qn("w:sectPr"))
    if existing is not None:
        body.remove(existing)
    if section is not None:
        body.append(deepcopy(section))
    document.save(output)
    _verify_b7(output, [migration], _xml_value(body), report["numbered_paragraphs"])
    report["styles_migrated"] = len(migration["styles"])
    report["numbering_definitions_migrated"] = len(migration["nums"])
    report["relationships_remapped"] = len(cache)
    return report


def _split_plan_value(source: Path, level: int) -> tuple[str, str, list[str], dict[str, int]]:
    """(inputs value, expected outputs value, part titles, feature totals) without writing."""
    _migrate_b7(Document(source), Document(), 1)
    parts = _split_parts(source, level)
    titles = [title for title, _ in parts]
    texts = []
    for _, blocks in parts:
        for block in blocks:
            text = getattr(block, "text", None)
            if text:
                texts.append(text)
    signature = _docx_signature(source)
    totals = {"relationships": signature["relationships"], "numbered": signature["numbered"]}
    return (json.dumps({"source": source.name, "sha256": digest(source), "level": level,
                        "heading_count": len(titles), "fidelity_policy": dict(_B7_POLICY), **totals},
                       ensure_ascii=False),
            json.dumps(_signature_value(texts, 0, 0), ensure_ascii=False),
            titles, totals)


def _check_excel_cell_address(address: str) -> None:
    if not isinstance(address, str) or not re.fullmatch(r"[A-Za-z]{1,3}[1-9][0-9]{0,6}", address):
        raise ValueError(f"invalid cell address {address!r}: expected a single A1-style cell like 'B1'")
    try:
        bounds = range_boundaries(address)
    except (TypeError, ValueError) as exc:
        raise ValueError(f"invalid cell address {address!r}: {exc}") from exc
    if not (1 <= bounds[0] <= 16384 and 1 <= bounds[1] <= 1048576):
        raise ValueError(f"invalid cell address {address!r}: outside Excel limits (A1:XFD1048576)")
    if bounds[0] != bounds[2] or bounds[1] != bounds[3]:
        raise ValueError(f"invalid cell address {address!r}: only one cell can be changed")


class WordTableRow(Strict):
    """Insert an empty row, or delete one row, of a simple merge-free table."""
    kind: str = "word_table_row"
    path: str
    table_index: int = Field(ge=0)
    row_index: int = Field(ge=0)
    action: Literal["insert_before", "insert_after", "delete"]


class ExcelSet(Strict):
    kind: str = "excel_set"
    path: str
    sheet: str = Field(min_length=1)
    cell: str
    value: str | int | float | bool | None

    @model_validator(mode="after")
    def check_cell(self):
        if isinstance(self.value, str) and self.value.startswith("="):
            raise ValueError("formula writes require a dedicated validated formula tool; ExcelSet accepts literal values only")
        if isinstance(self.value, float) and not math.isfinite(self.value):
            raise ValueError("non-finite numbers are not valid Excel literal values")
        _check_excel_cell_address(self.cell)
        return self


class ExcelTrim(Strict):
    kind: str = "excel_trim"
    path: str
    sheet: str = Field(min_length=1)
    cell: str
    expected_original: str = Field(min_length=1, max_length=500)

    @model_validator(mode="after")
    def check_cell(self):
        _check_excel_cell_address(self.cell)
        return self


class ExcelRangeTrim(Strict):
    """Trim surrounding whitespace in every literal text cell of one bounded rectangle."""
    kind: str = "excel_range_trim"
    path: str
    sheet: str = Field(min_length=1)
    first_cell: str
    last_cell: str

    @model_validator(mode="after")
    def check_range(self):
        _check_excel_cell_address(self.first_cell)
        _check_excel_cell_address(self.last_cell)
        left, top, right, bottom = range_boundaries(f"{self.first_cell}:{self.last_cell}")
        if left > right or top > bottom:
            raise ValueError(f"range must run from the top-left to the bottom-right cell, "
                             f"but {self.first_cell}:{self.last_cell} does not")
        if (right - left + 1) * (bottom - top + 1) > _MAX_TRIM_CELLS:
            raise ValueError(f"cleaning is bounded to {_MAX_TRIM_CELLS} cells per confirmed "
                             f"edit, but {self.first_cell}:{self.last_cell} covers more")
        return self


class ExcelRangeSet(Strict):
    """Write one bounded rectangular block of literal values, declared cell by cell.

    Every target cell is declared twice: ``expected`` states what the caller believes is
    there now (``None`` means "blank"), ``values`` states what it must become. The whole
    edit is refused unless every declared original matches the file exactly, so a shifted
    row can never be overwritten silently. Formulas and merged ranges are never targets.
    """
    kind: str = "excel_range_set"
    path: str
    sheet: str = Field(min_length=1)
    first_cell: str
    last_cell: str
    values: list[list[str | int | float | bool | None]]
    expected: list[list[str | int | float | bool | None]]

    @model_validator(mode="after")
    def check_block(self):
        _check_excel_cell_address(self.first_cell)
        _check_excel_cell_address(self.last_cell)
        left, top, right, bottom = range_boundaries(f"{self.first_cell}:{self.last_cell}")
        if left > right or top > bottom:
            raise ValueError(f"range must run from the top-left to the bottom-right cell, "
                             f"but {self.first_cell}:{self.last_cell} does not")
        height, width = bottom - top + 1, right - left + 1
        if height * width > _MAX_RANGE_SET_CELLS:
            raise ValueError(f"a block write is bounded to {_MAX_RANGE_SET_CELLS} cells per confirmed "
                             f"edit, but {self.first_cell}:{self.last_cell} covers {height * width}")
        for name, block in (("values", self.values), ("expected", self.expected)):
            if len(block) != height or any(len(row) != width for row in block):
                raise ValueError(f"{name} must be a {height}x{width} block matching "
                                 f"{self.first_cell}:{self.last_cell}")
        for row in self.values:
            for value in row:
                if isinstance(value, str):
                    if value.startswith("="):
                        raise ValueError("block writes accept literal values only; a formula needs "
                                         "the dedicated validated formula tools")
                    if len(value) > 500:
                        raise ValueError("literal text in a block write is limited to 500 characters")
                elif isinstance(value, float):
                    if not math.isfinite(value):
                        raise ValueError("non-finite numbers are not valid Excel literal values")
                    if value.is_integer():
                        # openpyxl writes a whole float as an integer, so it would come back a
                        # different type and fail the post-write verification. Say so up front
                        # instead of reporting a verification failure after the fact.
                        raise ValueError(f"write {int(value)} instead of {value!r}: Excel stores whole "
                                         f"numbers as integers, so {value!r} would not verify after the write")
        return self


class ExcelSort(Strict):
    """Stable sorting of a declared literal rectangle, without a header row.

    The range contains complete records; columns outside it are never moved. Every
    original must match ``expected`` and coordinate formatting is retained.
    """
    kind: str = "excel_sort"
    path: str
    sheet: str = Field(min_length=1)
    first_cell: str
    last_cell: str
    key_cell: str
    expected: list[list[str | int | float | bool | None]]
    descending: bool = False

    @model_validator(mode="after")
    def check_range(self):
        # Share literal/shape/reference bounds with the existing confirmed block writer.
        ExcelRangeSet(path=self.path, sheet=self.sheet, first_cell=self.first_cell,
                      last_cell=self.last_cell, values=self.expected, expected=self.expected)
        _check_excel_cell_address(self.key_cell)
        left, top, right, bottom = range_boundaries(f"{self.first_cell}:{self.last_cell}")
        key_column, key_row, _, _ = range_boundaries(self.key_cell)
        if bottom <= top or right - left >= 20 or not (left <= key_column <= right and key_row == top):
            raise ValueError("sort needs 2-100 rows, at most 20 columns/100 cells, and key_cell "
                             "must be a full A1 address in the first row; exclude the header")
        return self


class ExcelFormulaFill(Strict):
    """Copy a source formula into at most 20 empty cells directly below or beside it."""
    kind: str = "excel_formula_fill"
    path: str
    sheet: str = Field(min_length=1)
    source_cell: str
    target_range: str
    expected_source_formula: str = Field(min_length=2, max_length=250)
    direction: Literal["down", "right"] = "down"

    @model_validator(mode="after")
    def check_range(self):
        _check_excel_cell_address(self.source_cell)
        # A single-cell target such as "B2" is unambiguous and is normalised to the
        # equivalent one-cell range. Every semantic constraint below still applies; a
        # local 7B model otherwise burned its budget resending the bare address.
        if _CELL_ONLY.fullmatch(self.target_range or ""):
            self.target_range = f"{self.target_range}:{self.target_range}"
        if not _BOUNDED_RANGE.fullmatch(self.target_range or ""):
            # State the expected shape so a caller can retry correctly in one step.
            raise ValueError(f"formula fill needs an explicit bounded range like 'B2:B2' adjacent to "
                             f"source_cell (direction={self.direction}), not {self.target_range!r}")
        left, top, right, bottom = range_boundaries(self.target_range)
        source_col, source_row, _, _ = range_boundaries(self.source_cell)
        if self.direction == "down":
            adjacent = (left == right == source_col and top == source_row + 1 and
                        top <= bottom <= 1048576 and bottom - top < 20 and source_col <= 16384)
            where = f"1-20 blank cells immediately below {self.source_cell} in the same column"
        else:
            adjacent = (top == bottom == source_row and left == source_col + 1 and
                        left <= right <= 16384 and right - left < 20 and source_row <= 1048576)
            where = f"1-20 blank cells immediately to the right of {self.source_cell} in the same row"
        if not adjacent:
            raise ValueError(f"formula fill must cover {where}, but {self.target_range!r} does not")
        _safe_fill_formula(self.expected_source_formula)
        return self


_CHART_TYPES = {"bar": "BarChart", "line": "LineChart", "pie": "PieChart"}
_MAX_CHART_SERIES = 8
_LEGEND_POSITIONS = {"right": "r", "left": "l", "top": "t", "bottom": "b"}


def _check_chart_text(value: str | None, label: str) -> None:
    if value is not None and any(ord(char) < 32 and char not in "\t\n\r" for char in value):
        raise ValueError(f"{label} contains invalid XML control characters")


class ExcelChart(Strict):
    """One or more columns of literal values drawn as a bar, line or pie chart.

    Superset of :class:`ExcelBarChart`, which stays in place so existing plans and tools
    keep working. Every series column is validated independently; a pie chart takes exactly
    one series because Excel draws a single ring per pie chart.
    """
    kind: str = "excel_chart"
    path: str
    sheet: str = Field(min_length=1)
    chart_type: Literal["bar", "line", "pie"] = "bar"
    series_ranges: list[str] = Field(min_length=1, max_length=_MAX_CHART_SERIES)
    categories_range: str | None = None
    anchor: Literal["A1"] = "A1"
    title: str = Field(min_length=1, max_length=100)
    category_axis_title: str | None = Field(default=None, max_length=100)
    value_axis_title: str | None = Field(default=None, max_length=100)
    legend_position: Literal["right", "left", "top", "bottom", "none"] = "right"

    @model_validator(mode="after")
    def check_ranges(self):
        for label, text in (("chart title", self.title),
                            ("category axis title", self.category_axis_title),
                            ("value axis title", self.value_axis_title)):
            _check_chart_text(text, label)
        ranges = list(self.series_ranges)
        if self.chart_type == "pie" and len(ranges) != 1:
            raise ValueError("a pie chart takes exactly one data column; drop the extra series")
        if self.chart_type == "pie" and (self.category_axis_title or self.value_axis_title):
            raise ValueError("a pie chart has no axes, so axis titles are not accepted")
        shapes = [_chart_range(raw) for raw in ranges]
        columns = [col for col, _, _ in shapes]
        if len(set(columns)) != len(columns):
            raise ValueError("each series must be a different column")
        if len({(first, last) for _, first, last in shapes}) != 1:
            raise ValueError("every series must cover exactly the same rows")
        if self.categories_range is not None:
            category_col, first, last = _chart_range(self.categories_range)
            if (first, last) != shapes[0][1:] or category_col in columns:
                raise ValueError("category column must be distinct and cover exactly the value rows")
        return self


class ExcelBarChart(Strict):
    """One column of literal values, with an optional parallel category column."""
    kind: str = "excel_bar_chart"
    path: str
    sheet: str = Field(min_length=1)
    values_range: str
    categories_range: str | None = None
    anchor: Literal["A1"] = "A1"
    title: str = Field(min_length=1, max_length=100)

    @model_validator(mode="after")
    def check_ranges(self):
        if any(ord(char) < 32 and char not in "\t\n\r" for char in self.title):
            raise ValueError("chart title contains invalid XML control characters")
        values = _chart_range(self.values_range)
        if self.categories_range is not None:
            categories = _chart_range(self.categories_range)
            if values[1:] != categories[1:] or values[0] == categories[0]:
                raise ValueError("category column must be distinct and cover exactly the value rows")
        return self


Edit = (WordReplace | WordBatchReplace | WordFormat | WordTableCellReplace | WordTableRow
        | WordMergedCellText | WordRevisionDecision | DocxMerge | DocxSplit
        | LegacyWordReplace | LegacyCellSet
        | ExcelSet | ExcelTrim | ExcelRangeTrim
        | ExcelRangeSet | ExcelSort | ExcelFormulaFill | ExcelBarChart | ExcelChart)
_LEGACY_KINDS = (LegacyWordReplace, LegacyCellSet)
_MERGE_KINDS = (DocxMerge, DocxSplit)


# Rejections must tell the caller how to succeed: a local 7B model otherwise retried the
# same invented formula until the decision budget ran out.
_FORMULA_REMEDY = ("formula fill rejected {reason}; read the source cell with "
                   "read_excel_region and pass its exact formula text as expected_source_formula")
_CELL_ONLY = re.compile(r"[A-Za-z]{1,3}[1-9][0-9]{0,6}")
_BOUNDED_RANGE = re.compile(r"[A-Za-z]{1,3}[1-9][0-9]{0,6}:[A-Za-z]{1,3}[1-9][0-9]{0,6}")
# Windows resolves these names to devices, not files, even with an extension.
_RESERVED_DEVICE = re.compile(r"(?:CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])(?:\..*)?", re.IGNORECASE)
_MAX_TRIM_CELLS = 200  # Bound for one confirmed whitespace-cleaning edit.
_MAX_RANGE_SET_CELLS = 100  # Bound for one confirmed literal block write; both blocks are declared.
# A local model retried the same non-text cell until its budget ran out, so each rejection
# states what to do instead of only what was wrong.
_TRIM_REMEDY = ("trim rejected: {reason}; read the target with read_excel_region first, then clean "
                "literal text cells individually with plan_excel_trim or a bounded rectangle with "
                "plan_excel_range_trim")
# A block write is refused as a whole, so the rejection has to say how to make it match.
_RANGE_SET_REMEDY = ("block write rejected: {reason}; read the exact region with read_excel_region and pass "
                     "the current value of every cell as expected (null for a blank cell). Nothing was written")


_ALLOWED_FUNCTIONS = frozenset({
    # Aggregates
    "SUM", "AVERAGE", "MIN", "MAX", "COUNT", "COUNTA", "PRODUCT",
    # Logic
    "IF", "IFERROR", "AND", "OR", "NOT",
    # Deterministic maths
    "ROUND", "ROUNDUP", "ROUNDDOWN", "INT", "TRUNC", "ABS", "SIGN", "MOD",
    "POWER", "SQRT", "EXP", "LN", "LOG10", "CEILING", "FLOOR",
})
# Deterministic but position- or context-dependent: a filled-down formula would silently
# change meaning, or the reference cannot be validated statically, so they stay out.
_VOLATILE_FUNCTIONS = frozenset({
    "NOW", "TODAY", "RAND", "RANDBETWEEN", "OFFSET", "INDIRECT", "CELL", "INFO",
    "ROW", "COLUMN", "ROWS", "COLUMNS",
})
# These can reach outside the workbook or the local machine.
_EXTERNAL_FUNCTIONS = frozenset({"HYPERLINK", "WEBSERVICE", "RTD", "FILTERXML", "IMAGE"})
_SINGLE_CELL = re.compile(r"\$?[A-Z]{1,3}\$?[1-9][0-9]{0,6}", re.I)
_CELL_RANGE = re.compile(r"(\$?[A-Z]{1,3}\$?[1-9][0-9]{0,6}):(\$?[A-Z]{1,3}\$?[1-9][0-9]{0,6})", re.I)
_NUMERIC_LITERAL = re.compile(r"[0-9]+(?:\.[0-9]+)?")
_TEXT_LITERAL = re.compile(r'"(?:[^"]|"")*"')
_INFIX = {"+", "-", "*", "/", "^", "&", "=", "<", ">", "<=", ">=", "<>"}


def _check_formula_reference(value: str) -> None:
    """A reference must stay inside this sheet and inside Excel's grid."""
    if "!" in value or "[" in value or "]" in value:
        raise ValueError(_FORMULA_REMEDY.format(
            reason="only references on the same sheet are allowed, so cross-sheet or "
                   "external references are rejected"))
    single = value.replace("$", "")
    if "/" in value or value.count(":") > 1:
        raise ValueError(_FORMULA_REMEDY.format(
            reason=f"{value!r} is not a bounded cell or cell-range reference"))
    if _SINGLE_CELL.fullmatch(value):
        col, row, _, _ = range_boundaries(single)
    else:
        match = _CELL_RANGE.fullmatch(value)
        if match is None:
            raise ValueError(_FORMULA_REMEDY.format(
                reason=f"{value!r} is not a local cell or cell-range reference; whole-column, "
                       f"whole-row and named ranges are not accepted"))
        start, end = match.group(1).replace("$", ""), match.group(2).replace("$", "")
        left, top, right, bottom = range_boundaries(f"{start}:{end}")
        if left > right or top > bottom:
            raise ValueError(_FORMULA_REMEDY.format(
                reason=f"the range {value!r} runs backwards"))
        col, row = right, bottom
        if left < 1 or top < 1:
            raise ValueError(_FORMULA_REMEDY.format(reason="a reference exceeds Excel limits"))
    if not (1 <= col <= 16384 and 1 <= row <= 1048576):
        raise ValueError(_FORMULA_REMEDY.format(reason="a reference exceeds Excel limits"))


def _safe_fill_formula(formula: str) -> None:
    """Allow local references, arithmetic, comparisons and an allowlist of functions.

    Functions are accepted only from a fixed list of deterministic, dependency-free names;
    dynamic-reference and volatile functions are rejected by name so the reason is specific.
    Anything the tokenizer emits that this grammar does not explicitly permit is rejected.
    """
    if not formula.startswith("=") or len(formula) > 250 or any(ord(ch) < 32 for ch in formula):
        raise ValueError(_FORMULA_REMEDY.format(reason="expected an existing = formula of at most 250 characters"))
    try:
        tokens = Tokenizer(formula).items
    except (TokenizerError, ValueError, IndexError) as exc:
        # openpyxl 3.1.5 raises IndexError for an unmatched closing parenthesis.
        raise ValueError(_FORMULA_REMEDY.format(reason="the formula could not be parsed")) from exc
    if not tokens:
        raise ValueError(_FORMULA_REMEDY.format(reason="the formula is empty"))
    needs_operand = True
    depth = 0
    for token in tokens:
        if token.type == "WHITE-SPACE":
            # Ordinary spacing is inert; the operand/operator grammar below rejects
            # Excel's implicit-intersection use of whitespace between references.
            continue
        if token.type == "FUNC" and token.subtype == "OPEN":
            name = token.value.rstrip("(").strip().upper()
            # Every rejection names the offending function, so a caller can fix the call.
            if name in _VOLATILE_FUNCTIONS:
                raise ValueError(_FORMULA_REMEDY.format(
                    reason=f"the function {name} is volatile or position-dependent, so filling it "
                           f"down or right would silently change its meaning"))
            if name in _EXTERNAL_FUNCTIONS:
                raise ValueError(_FORMULA_REMEDY.format(
                    reason=f"the function {name} can reach outside the workbook, so it is not allowed "
                           f"in a filled formula"))
            if name not in _ALLOWED_FUNCTIONS:
                raise ValueError(_FORMULA_REMEDY.format(
                    reason=f"the function {name} is not in the allowed list "
                           f"({', '.join(sorted(_ALLOWED_FUNCTIONS))})"))
            if not needs_operand:
                raise ValueError(_FORMULA_REMEDY.format(
                    reason=f"{name} is not preceded by an operator"))
            depth += 1
            needs_operand = True  # the first argument must follow
        elif token.type == "FUNC" and token.subtype == "CLOSE":
            if depth == 0 or needs_operand:
                raise ValueError(_FORMULA_REMEDY.format(reason="a function call has no complete argument"))
            depth -= 1
            needs_operand = False
        elif token.type == "SEP" and token.subtype == "ARG":
            if depth == 0 or needs_operand:
                raise ValueError(_FORMULA_REMEDY.format(
                    reason="a comma must separate two complete function arguments"))
            needs_operand = True
        elif token.type == "OPERAND" and needs_operand:
            if token.subtype == "RANGE":
                # A bare range as the whole formula (e.g. "=A1:A2") is an array or implicit
                # intersection, not arithmetic; ranges are arguments of a function only.
                if ":" in token.value and depth == 0:
                    raise ValueError(_FORMULA_REMEDY.format(
                        reason="a range is only accepted as an argument inside an allowed function, "
                               "never as the whole formula"))
                _check_formula_reference(token.value)
            elif token.subtype == "NUMBER":
                if not _NUMERIC_LITERAL.fullmatch(token.value):
                    raise ValueError(_FORMULA_REMEDY.format(reason=f"{token.value!r} is not a plain number"))
            elif token.subtype == "TEXT":
                if not _TEXT_LITERAL.fullmatch(token.value):
                    raise ValueError(_FORMULA_REMEDY.format(reason="a text literal is malformed"))
            elif token.subtype != "LOGICAL" or token.value.upper() not in {"TRUE", "FALSE"}:
                raise ValueError(_FORMULA_REMEDY.format(
                    reason="only local references, numbers, text literals, TRUE/FALSE, arithmetic, "
                           "comparisons, parentheses and allowlisted functions are accepted"))
            needs_operand = False
        elif token.type == "OPERATOR-INFIX" and not needs_operand and token.value in _INFIX:
            needs_operand = True
        elif token.type == "OPERATOR-PREFIX" and needs_operand and token.value in {"+", "-"}:
            pass
        elif token.type == "OPERATOR-POSTFIX" and not needs_operand and token.value == "%":
            pass
        elif token.type == "PAREN" and token.subtype == "OPEN" and needs_operand:
            depth += 1
        elif token.type == "PAREN" and token.subtype == "CLOSE" and not needs_operand and depth:
            depth -= 1
        else:
            raise ValueError(_FORMULA_REMEDY.format(
                reason="only local references, numbers, text literals, TRUE/FALSE, arithmetic, "
                       "comparisons, parentheses and allowlisted functions are accepted"))
    if needs_operand or depth:
        raise ValueError(_FORMULA_REMEDY.format(reason="the expression is incomplete"))


def _formula_fill_values(ws, request: ExcelFormulaFill) -> dict[str, str]:
    source = ws[request.source_cell]
    if source.data_type != "f" or source.value != request.expected_source_formula:
        raise GuardError("source formula changed or does not match expected_source_formula")
    _safe_fill_formula(source.value)
    left, top, right, bottom = range_boundaries(request.target_range)
    values = {}
    if request.direction == "down":
        targets = [ws.cell(row, left) for row in range(top, bottom + 1)]
    else:
        targets = [ws.cell(top, column) for column in range(left, right + 1)]
    for cell in targets:
        if cell.value is not None or cell.__class__.__name__ == "MergedCell" or cell.coordinate in ws.merged_cells:
            raise GuardError(f"formula fill requires empty, unmerged destination cells, but "
                             f"{cell.coordinate} is not empty (never overwriting existing content)")
        try:
            translated = Translator(source.value, origin=source.coordinate).translate_formula(cell.coordinate)
            _safe_fill_formula(translated)
        except (TranslatorError, ValueError) as exc:
            raise GuardError("formula translation produced an unsafe or out-of-range reference") from exc
        values[cell.coordinate] = translated
    return values


def _chart_range(raw: str) -> tuple[int, int, int]:
    if not re.fullmatch(r"[A-Z]{1,3}[1-9][0-9]{0,6}:[A-Z]{1,3}[1-9][0-9]{0,6}", raw):
        raise ValueError("chart source must be an explicit single-column range")
    min_col, min_row, max_col, max_row = range_boundaries(raw)
    if not (1 <= min_col == max_col <= 16384 and
            1 <= min_row <= max_row <= 1048576 and max_row - min_row + 1 <= 20 and
            raw == f"{get_column_letter(min_col)}{min_row}:{get_column_letter(max_col)}{max_row}"):
        raise ValueError("chart source must be one bounded column of at most 20 rows")
    return min_col, min_row, max_row


def _chart_sources(ws, request) -> tuple[list[Reference], Reference | None]:
    """Validate every series column and return references for the chart builder.

    Accepts both chart requests. Returns a list because the general chart may carry several
    series; the single-series request simply yields a one-element list.
    """
    ranges = (list(request.series_ranges) if isinstance(request, ExcelChart)
              else [request.values_range])
    references = []
    for raw in ranges:
        col, first, last = _chart_range(raw)
        for row in range(first, last + 1):
            cell = ws.cell(row, col)
            try:
                numeric = (cell.data_type == "n" and not isinstance(cell.value, bool) and
                           isinstance(cell.value, (int, float)) and math.isfinite(cell.value))
            except OverflowError:
                numeric = False
            if not numeric:
                raise GuardError(f"chart values must be finite numeric literal cells without "
                                 f"formulas or blanks, but {cell.coordinate} is not")
        references.append(Reference(ws, min_col=col, min_row=first, max_row=last))
    first, last = references[0].min_row, references[0].max_row
    categories = None
    if getattr(request, "categories_range", None) is not None:
        category_col, _, _ = _chart_range(request.categories_range)
        for row in range(first, last + 1):
            cell = ws.cell(row, category_col)
            if cell.data_type == "f" or cell.value is None or isinstance(cell.value, bool) or not isinstance(cell.value, (str, int, float)) or (isinstance(cell.value, float) and not math.isfinite(cell.value)):
                raise GuardError("chart categories must be nonempty literal cells without formulas")
        categories = Reference(ws, min_col=category_col, min_row=first, max_row=last)
    return references, categories


def _reference_formula(reference: Reference) -> str:
    """The A1 formula openpyxl stores in the chart, e.g. ``'Data'!$B$1:$B$2``."""
    return str(reference)


def _build_chart(request, references: list[Reference], categories: Reference | None):
    """Create the openpyxl chart object for either chart request (shared by both tools)."""
    chart_type = getattr(request, "chart_type", "bar")
    if chart_type == "line":
        chart = LineChart()
    elif chart_type == "pie":
        chart = PieChart()
    else:
        chart = BarChart()
        chart.type = "col"
        chart.grouping = "clustered"
    chart.title = request.title
    for reference in references:
        chart.add_data(reference, titles_from_data=False)
    if categories is not None:
        chart.set_categories(categories)
    if chart_type != "pie":
        # Axes only exist on bar/line charts; the request validator rejects them for pie.
        if getattr(request, "category_axis_title", None):
            chart.x_axis.title = request.category_axis_title
        if getattr(request, "value_axis_title", None):
            chart.y_axis.title = request.value_axis_title
    position = getattr(request, "legend_position", "right")
    if position == "none":
        chart.legend = None
    else:
        chart.legend.position = _LEGEND_POSITIONS[position]
    return chart


def _chart_projection(chart) -> tuple:
    """Semantics of the generated chart, not a count-only drawing check."""
    expected_names = set(_CHART_TYPES.values())
    if type(chart).__name__ not in expected_names or not chart.series:
        raise GuardError("chart type or series changed")
    title = chart.title
    paragraphs = title.tx.rich.p if title is not None and title.tx and title.tx.rich else []
    text = "".join(run.t for p in paragraphs for run in p.r)

    def axis_title(axis) -> str | None:
        node = getattr(axis, "title", None)
        rich = getattr(node, "tx", None)
        rich = getattr(rich, "rich", None) if rich is not None else None
        if not rich:
            return None
        return "".join(run.t for para in rich.p for run in para.r) or None

    series = []
    for item in chart.series:
        values = item.val.numRef.f if item.val and item.val.numRef else None
        categories = (item.cat.numRef.f if item.cat and item.cat.numRef else
                      item.cat.strRef.f if item.cat and item.cat.strRef else None)
        series.append((values, categories))
    legend = getattr(chart.legend, "position", None) if chart.legend is not None else None
    anchor = chart.anchor
    position = (anchor._from.col, anchor._from.row) if hasattr(anchor, "_from") else anchor
    axes = (axis_title(chart.x_axis), axis_title(chart.y_axis)) if hasattr(chart, "x_axis") else (None, None)
    return (type(chart).__name__, getattr(chart, "type", None), getattr(chart, "grouping", None),
            text, tuple(series), axes, legend, position)


def _normalized_projection(chart) -> tuple:
    """Projection with the anchor made comparable between memory and a reopened file.

    A freshly built chart still holds the anchor string it was added with, while a
    reopened one exposes the resolved cell offset; both describe the same cell, and only
    ``A1`` is ever allowed.
    """
    projection = _chart_projection(chart)
    position = projection[-1]
    if not isinstance(position, tuple):
        position = (0, 0) if position == "A1" else position
    return (*projection[:-1], position)


def _chart_package_safe(path: Path, operation: str = "chart creation") -> None:
    """Only ordinary, small one-sheet OOXML parts; unknown parts cannot round-trip safely."""
    allowed = {"[Content_Types].xml", "_rels/.rels", "docProps/app.xml", "docProps/core.xml",
               "xl/workbook.xml", "xl/_rels/workbook.xml.rels", "xl/worksheets/sheet1.xml",
               "xl/styles.xml", "xl/theme/theme1.xml", "xl/sharedStrings.xml"}
    with zipfile.ZipFile(path) as archive:
        names = archive.namelist()
        if len(names) != len(set(names)):
            raise GuardError(f"{operation} requires simple OOXML without duplicate parts")
        # A second worksheet must be rejected before the generic unknown-part
        # check, so callers receive the precise one-sheet constraint.
        if any(re.fullmatch(r"xl/worksheets/sheet\d+\.xml", name) and
               name != "xl/worksheets/sheet1.xml" for name in names):
            raise GuardError(f"{operation} requires one sheet")
        if not set(names).issubset(allowed):
            raise GuardError(f"{operation} requires simple OOXML without existing charts/drawings or unknown parts")
        ns = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
        permitted = {
            "xl/worksheets/sheet1.xml": {"sheetPr", "dimension", "sheetViews", "sheetFormatPr",
                                          "cols", "sheetData", "mergeCells", "pageMargins"},
            "xl/workbook.xml": {"fileVersion", "fileSharing", "workbookPr", "bookViews",
                                "sheets", "calcPr", "workbookProtection", "definedNames"},
        }
        for part, tags in permitted.items():
            root = ET.fromstring(archive.read(part))
            expected_root = ns + ("workbook" if part == "xl/workbook.xml" else "worksheet")
            if root.tag != expected_root:
                raise GuardError(f"{operation} refuses unexpected XML root: {part}: {root.tag}")
            allowed_tags = {ns + tag for tag in tags}
            for child in root:
                if child.tag not in allowed_tags:
                    raise GuardError(f"{operation} refuses unsupported workbook or worksheet semantics: {part}: {child.tag}")
                # openpyxl emits an inert workbookProtection tag; active protection
                # is blocked by inspect_file, and unfamiliar attributes fail closed.
                if part == "xl/workbook.xml" and child.tag == ns + "workbookProtection" and (
                        len(child) or (child.text or "").strip() or any(
                            key not in {"lockStructure", "lockWindows", "lockRevision"}
                            or value.strip().lower() not in {"0", "false", "off"}
                            for key, value in child.attrib.items())):
                    raise GuardError(f"{operation} refuses unsupported workbook protection semantics")
                # openpyxl emits <definedNames/> for an ordinary workbook. Named ranges,
                # print areas and other nonempty definitions may not round-trip safely.
                if part == "xl/workbook.xml" and child.tag == ns + "definedNames" and (
                        child.attrib or len(child) or (child.text or "").strip()):
                    raise GuardError(f"{operation} refuses nonempty workbook definedNames")



def digest(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            h.update(block)
    return h.hexdigest()


class Workspace:
    """A directory granted by the human at startup, never chosen by the model."""

    def __init__(self, root: Path):
        self.root = root.resolve(strict=True)
        if not self.root.is_dir() or not self.root.is_absolute():
            raise GuardError("authorized root must be an existing absolute directory")

    def file(self, raw: str, *, existing: bool = True) -> Path:
        if not isinstance(raw, str) or not raw or ".." in Path(raw).parts:
            raise GuardError("relative/traversal path rejected")
        # \\?\ and \\.\ address the device namespace, where containment checks and the
        # "exactly this file" comparison no longer describe the same object.
        if raw.startswith("\\\\") and raw[2:3] in {"?", "."}:
            raise GuardError("device-namespace paths are not accepted")
        path = Path(raw)
        if not path.is_absolute():
            raise GuardError("an absolute path is required")
        for part in path.parts[1:]:
            # Windows strips trailing dots and spaces, so "doc.xlsx " silently addresses
            # "doc.xlsx"; the caller's path must be the path actually used.
            if part != part.rstrip(" ."):
                raise GuardError("trailing dot or space in a path component rejected")
            if _RESERVED_DEVICE.fullmatch(part):
                raise GuardError(f"reserved Windows device name rejected: {part}")
        try:
            resolved = path.resolve(strict=existing)
        except OSError as exc:
            raise GuardError(f"path could not be resolved safely: {exc}") from exc
        if not resolved.is_relative_to(self.root) or resolved == self.root:
            raise GuardError("file outside authorized directory")
        if existing and (not resolved.is_file() or path.is_symlink()):
            raise GuardError("ordinary file required")
        if not existing and (resolved.exists() or path.is_symlink()):
            raise GuardError("output already exists")
        return resolved


def _batch_replace(doc, request: WordBatchReplace) -> list[dict[str, Any]]:
    """Apply a document-wide literal replacement and describe every paragraph it changed.

    Fails closed rather than partially applying: an occurrence count above the caller's cap
    is refused outright, because a silently truncated bulk edit is worse than no edit.
    """
    hits = [(index, paragraph) for index, paragraph in enumerate(doc.paragraphs)
            if request.old in paragraph.text]
    total = sum(paragraph.text.count(request.old) for _, paragraph in hits)
    if total == 0:
        raise GuardError("no occurrence of the target text was found in the body")
    if total > request.max_occurrences:
        raise GuardError(f"found {total} occurrences, above the {request.max_occurrences} "
                         f"allowed for one confirmed batch edit")
    targets = []
    for index, paragraph in hits:
        before = paragraph.text
        expected = before.replace(request.old, request.new)
        if _replace_runs(paragraph, request.old, request.new) != before.count(request.old):
            raise GuardError("target text includes hyperlink/unsupported run content")
        if paragraph.text != expected:
            raise GuardError("paragraph text cannot be replaced without changing unsupported structures")
        targets.append({"paragraph_index": index, "before": before, "after": expected})
    return targets


def _simple_table_row(doc, request: WordTableRow):
    """Resolve a table row that can be restructured without corrupting the table.

    Nested tables and merged cells (``gridSpan``/``vMerge``) are refused outright: adding
    or removing a row inside those changes how the existing grid is interpreted.
    """
    if request.table_index >= len(doc.tables):
        raise GuardError(f"table index {request.table_index} is out of bounds")
    table = doc.tables[request.table_index]
    rows = table.rows
    if not rows:
        raise GuardError("table has no rows")
    if request.row_index >= len(rows):
        raise GuardError(f"row index {request.row_index} is out of bounds for a table with {len(rows)} rows")
    if request.action == "delete" and len(rows) == 1:
        raise GuardError("cannot delete the only row of a table")
    for row in rows:
        for cell in row.cells:
            if cell.tables:
                raise GuardError("table contains a nested table; restructuring is refused")
            properties = cell._tc.tcPr
            if properties is not None and (properties.find(qn("w:gridSpan")) is not None or
                                           properties.find(qn("w:vMerge")) is not None):
                raise GuardError("table uses horizontally or vertically merged cells; "
                                 "restructuring is refused")
    return table, rows


def _table_row_summary(doc, request: WordTableRow, *, describe_deleted: bool = False) -> dict[str, Any]:
    """Structural summary used for the plan preview and the post-write check.

    ``describe_deleted`` is only set for the pre-edit state. Without it the post-edit
    summary would look up the same index in the shortened table and report a *different*
    row as the one being removed, which would make the preview lie about the edit.
    """
    if request.table_index >= len(doc.tables):
        raise GuardError(f"table index {request.table_index} is out of bounds")
    table = doc.tables[request.table_index]
    rows = table.rows
    summary: dict[str, Any] = {"table_index": request.table_index, "rows": len(rows),
                               "columns": len(rows[0].cells) if rows else 0}
    if describe_deleted and request.action == "delete" and request.row_index < len(rows):
        summary["deleted_row_content"] = [cell.text for cell in rows[request.row_index].cells]
    return summary


def _apply_table_row(doc, request: WordTableRow) -> None:
    table, rows = _simple_table_row(doc, request)
    if request.action == "delete":
        table._tbl.remove(rows[request.row_index]._tr)
        return
    # add_row() appends a fresh empty row; move that element to the requested position.
    new_row = table.add_row()
    target = rows[request.row_index]._tr
    if request.action == "insert_before":
        target.addprevious(new_row._tr)
    else:
        target.addnext(new_row._tr)


def _defined_header_footer(section: Any, kind: str):
    """The header/footer container only when the section actually defines one.

    Reading ``section.header`` materialises an empty part when the section has none, which
    would add ``word/header1.xml``/``word/footer1.xml`` to every output copy and can turn a
    "same as previous" section into an explicit empty header. Sections without a definition
    contribute nothing to the projection instead of being silently created.
    """
    part = section.header if kind == "header" else section.footer
    return None if part.is_linked_to_previous else part


def _merge_projection(doc) -> tuple:
    """Every table's grid columns and per-cell gridSpan/vMerge, as one comparable value.

    A write into a merged table has to prove it changed no merge, so this projection is
    taken before the edit and compared with the reopened output: a dropped gridSpan or a
    reshaped vMerge fails the write instead of silently reshaping the table.
    """
    def table_projection(tbl) -> tuple:
        grid = tbl.find(qn("w:tblGrid"))
        columns = (tuple(column.get(qn("w:w")) for column in grid.findall(qn("w:gridCol")))
                   if grid is not None else ())
        rows = []
        for tr in tbl.findall(qn("w:tr")):
            cells = []
            for tc in tr.findall(qn("w:tc")):
                properties = tc.find(qn("w:tcPr"))
                span = properties.find(qn("w:gridSpan")) if properties is not None else None
                merge = properties.find(qn("w:vMerge")) if properties is not None else None
                vertical = None
                if merge is not None:
                    vertical = ("restart" if (merge.get(qn("w:val")) or "").strip().lower() == "restart"
                                else "continue")
                cells.append((span.get(qn("w:val")) if span is not None else None, vertical))
            rows.append(tuple(cells))
        return (columns, tuple(rows))

    tables = [table_projection(table._tbl) for table in doc.tables]
    for section in doc.sections:
        for kind in ("header", "footer"):
            part = _defined_header_footer(section, kind)
            if part is not None:
                tables.extend(table_projection(table._tbl) for table in part.tables)
    return tuple(tables)


def _merged_cell(doc, request: WordMergedCellText):
    """Resolve the logical coordinate to the cell that carries its text, or refuse."""
    if request.table_index >= len(doc.tables):
        raise GuardError(f"table index {request.table_index} is out of bounds")
    table = doc.tables[request.table_index]
    # Imported lazily: documents.py imports nothing from this module at import time.
    from .documents import locate_word_cell, word_table_merge_issues
    issues = word_table_merge_issues(table._tbl)
    if issues:
        raise GuardError("表格存在结构不一致（" + "；".join(issues[:3]) +
                         "），拒绝在合并表格中写入；请先用 read_word_table_structure 查看")
    try:
        tc_element, info = locate_word_cell(table._tbl, request.row_index, request.column_index)
    except ValueError as exc:
        raise GuardError(str(exc)) from exc
    return table, _Cell(tc_element, table), info


def _merged_cell_paragraph(cell, request: WordMergedCellText):
    if request.paragraph_index >= len(cell.paragraphs):
        raise GuardError(f"单元格段落索引 {request.paragraph_index} 越界：该单元格有 "
                         f"{len(cell.paragraphs)} 个段落")
    paragraph = cell.paragraphs[request.paragraph_index]
    _plain_paragraph(paragraph)
    return paragraph


def _merged_cell_replace(cell, request: WordMergedCellText) -> tuple[str, str]:
    """Replace a literal in one paragraph of the resolved cell, reusing the run-preserving path."""
    paragraph = _merged_cell_paragraph(cell, request)
    before = paragraph.text
    count = before.count(request.old)
    if count != request.expected_count:
        raise GuardError(f"期望在目标单元格段落中找到 {request.expected_count} 处 "
                         f"{request.old!r}，实际找到 {count} 处")
    if _replace_runs(paragraph, request.old, request.new) != count:
        raise GuardError("目标单元格文本包含超链接或无法保留的 run 结构")
    after = before.replace(request.old, request.new)
    if paragraph.text != after:
        raise GuardError("目标单元格文本无法在不改变不支持结构的前提下替换")
    return before, after


def _merged_cell_append(cell, request: WordMergedCellText) -> tuple[str, str]:
    """Add one paragraph at the end of the cell, inheriting the last paragraph's formatting."""
    before = cell.text
    source = cell.paragraphs[-1] if cell.paragraphs else None
    paragraph = cell.add_paragraph()
    if source is not None and source._p.pPr is not None:
        paragraph._p.insert(0, deepcopy(source._p.pPr))
    run = paragraph.add_run(request.new)
    if source is not None and source.runs and source.runs[-1]._r.rPr is not None:
        run._r.insert(0, deepcopy(source.runs[-1]._r.rPr))
    after = cell.text
    if after == before:
        raise GuardError("追加文本没有改变该单元格")
    return before, after


# Revision marks that reshape a document rather than replacing text. Refusing them is what
# keeps "accept/reject all" from leaving a half-decided file behind.
_REVISION_UNSUPPORTED_MARKS = frozenset({
    "moveFrom", "moveTo", "moveFromRangeStart", "moveFromRangeEnd", "moveToRangeStart",
    "moveToRangeEnd", "cellIns", "cellDel", "cellMerge", "tblPrChange", "tblGridChange",
    "tcPrChange", "trPrChange", "sectPrChange", "numberingChange",
    "customXmlInsRangeStart", "customXmlInsRangeEnd", "customXmlDelRangeStart",
    "customXmlDelRangeEnd", "customXmlMoveFromRangeStart", "customXmlMoveFromRangeEnd",
    "customXmlMoveToRangeStart", "customXmlMoveToRangeEnd",
})
_REVISION_FORMAT_MARKS = ("rPrChange", "pPrChange")
# w:ins/w:del below one of these parents mark row, paragraph-mark, table or section
# revisions; they change structure, so they are reported instead of being unwrapped.
_REVISION_PROPERTY_CONTEXT = frozenset({"trPr", "rPr", "pPr", "tblPr", "tblGrid", "tcPr",
                                        "sectPr", "numPr"})
_REVISION_ALL_MARKS = (frozenset({"ins", "del"}) | set(_REVISION_FORMAT_MARKS)
                       | set(_REVISION_UNSUPPORTED_MARKS))
_REVISION_PART_BYTES = 20_000_000  # Parts above this are refused, never silently skipped.
# The inspector flags these as advanced XML; deciding revisions is exactly the operation
# that removes them, so only these two are tolerated (and only when nothing else is flagged).
_REVISION_TOLERATED_FEATURES = frozenset({"word_advanced_xml:<w:ins", "word_advanced_xml:<w:del"})
_REVISION_REMEDY = ("修订处理被拒绝：{reason}；为避免留下半成品，本次未做任何改动。"
                    "可先用 read_word_structure 查看逐条修订，或在 Word 中手工处理这些标记")


def _local_name(element: Any) -> str:
    return element.tag.rsplit("}", 1)[-1]


def _element_depth(element: Any) -> int:
    depth, parent = 0, element.getparent()
    while parent is not None:
        depth += 1
        parent = parent.getparent()
    return depth


def _revision_context(element: Any) -> str | None:
    """The nearest property element above a w:ins/w:del, or None for a content revision."""
    parent = element.getparent()
    while parent is not None:
        name = _local_name(parent)
        if name in _REVISION_PROPERTY_CONTEXT:
            return name
        parent = parent.getparent()
    return None


def _revision_inventory(body: Any) -> dict[str, Any]:
    """Supported content revisions, formatting revisions, and everything refused by name."""
    insertions: list[Any] = []
    deletions: list[Any] = []
    formats: list[Any] = []
    unsupported: dict[str, int] = {}
    for element in body.iter():
        name = _local_name(element)
        if name in {"ins", "del"}:
            context = _revision_context(element)
            if context is not None:
                key = f"{name}@{context}"
                unsupported[key] = unsupported.get(key, 0) + 1
            elif name == "ins":
                insertions.append(element)
            else:
                deletions.append(element)
        elif name in _REVISION_FORMAT_MARKS:
            formats.append(element)
        elif name in _REVISION_UNSUPPORTED_MARKS:
            unsupported[name] = unsupported.get(name, 0) + 1
    return {"insertions": insertions, "deletions": deletions, "format_changes": formats,
            "unsupported": unsupported}


def _unwrap(element: Any) -> None:
    """Replace an element with its children, keeping document order."""
    parent = element.getparent()
    if parent is None:
        return
    index = parent.index(element)
    for child in reversed(list(element)):
        parent.insert(index, child)
    parent.remove(element)


def _restore_deleted_text(element: Any) -> None:
    """Rejecting a deletion puts the deleted text back as ordinary text."""
    for node in element.iter():
        if node.tag == qn("w:delText"):
            node.tag = qn("w:t")
        elif node.tag == qn("w:delInstrText"):
            node.tag = qn("w:instrText")


def _reject_format_change(change: Any) -> None:
    """Rejecting a formatting revision restores the properties it recorded as previous."""
    recorded = change.find(qn("w:rPr") if _local_name(change) == "rPrChange" else qn("w:pPr"))
    parent = change.getparent()
    if parent is None:
        return
    restored = [deepcopy(child) for child in (recorded if recorded is not None else [])]
    parent.remove(change)
    for child in list(parent):
        parent.remove(child)
    for child in restored:
        parent.append(child)


def _apply_revision_decision(doc, action: str) -> None:
    """Apply accept-all or reject-all to content revisions, then settle format revisions.

    Deeper elements are processed first so a revision nested inside another one is decided
    before its container disappears. Accepted deletions vanish with their content; rejected
    deletions keep the content and their ``w:delText`` becomes ``w:t`` again.
    """
    body = doc.element.body
    inventory = _revision_inventory(body)
    targets = sorted(inventory["insertions"] + inventory["deletions"],
                     key=_element_depth, reverse=True)
    for element in targets:
        parent = element.getparent()
        if parent is None:  # already removed with an enclosing revision
            continue
        inserting = _local_name(element) == "ins"
        if action == "accept_all":
            if inserting:
                _unwrap(element)
            else:
                parent.remove(element)
        elif inserting:
            parent.remove(element)
        else:
            _restore_deleted_text(element)
            _unwrap(element)
    fresh = [item for item in body.iter() if _local_name(item) in _REVISION_FORMAT_MARKS]
    for change in fresh:
        if action == "accept_all":
            parent = change.getparent()
            if parent is not None:
                parent.remove(change)  # accepting keeps the current formatting
        else:
            _reject_format_change(change)


def _word_text_projection(doc) -> tuple[str, ...]:
    """Body paragraph and table-cell text as python-docx renders it.

    Inserted runs live inside ``w:ins`` and deleted runs inside ``w:del``, so both are
    invisible here until a decision unwraps them: accepting surfaces the inserted text and
    rejecting surfaces the deleted text again.
    """
    texts = [paragraph.text for paragraph in doc.paragraphs]
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                texts.append(cell.text)
    return tuple(texts)


def _text_digest(texts: tuple[str, ...]) -> str:
    return hashlib.sha256(json.dumps(list(texts), ensure_ascii=False).encode("utf-8")).hexdigest()


def _revision_value(texts: tuple[str, ...], changed: list[dict[str, Any]], truncated: bool) -> str:
    return json.dumps({"parts": len(texts), "text_sha256": _text_digest(texts),
                       "changed": changed, "changed_truncated": truncated}, ensure_ascii=False)


def _revision_apply_and_measure(doc, action: str) -> dict[str, Any]:
    """Apply the decision and return the bounded before/after preview plus the change list."""
    before_texts = _word_text_projection(doc)
    _apply_revision_decision(doc, action)
    after_texts = _word_text_projection(doc)
    if len(before_texts) == len(after_texts):
        changed = [{"index": index, "before": was, "after": now}
                   for index, (was, now) in enumerate(zip(before_texts, after_texts)) if was != now]
    else:
        # A paragraph-level insertion or deletion changes how many text units exist.
        changed = [{"index": None, "before": f"{len(before_texts)} 个文本单元",
                    "after": f"{len(after_texts)} 个文本单元",
                    "note": "修订包含段落级插入或删除，文本单元数量发生变化"}]
    return {"before": _revision_value(before_texts, [], False),
            "after": _revision_value(after_texts, changed[:20], len(changed) > 20),
            "changed": changed[:20], "changed_truncated": len(changed) > 20}


def _package_revision_marks(path: Path, *, include_document: bool = True) -> dict[str, int]:
    """Revision marks per package part, so a decision can prove it left none behind."""
    marks: dict[str, int] = {}
    try:
        with zipfile.ZipFile(path) as archive:
            names = archive.namelist()
            for part in names:
                if not (include_document and part == "word/document.xml") and not re.fullmatch(
                        r"word/(?:header|footer)\d*\.xml", part) and part not in {
                        "word/footnotes.xml", "word/endnotes.xml"}:
                    continue
                if archive.getinfo(part).file_size > _REVISION_PART_BYTES:
                    raise GuardError(f"部件 {part} 过大，未扫描，不能据此认为没有修订")
                try:
                    root = ET.fromstring(archive.read(part))
                except ET.ParseError as exc:
                    raise GuardError(f"部件 {part} 无法解析，不能据此认为没有修订") from exc
                for element in root.iter():
                    name = _local_name(element)
                    if name in _REVISION_ALL_MARKS:
                        key = f"{part}:{name}"
                        marks[key] = marks.get(key, 0) + 1
    except (zipfile.BadZipFile, OSError, KeyError) as exc:
        raise GuardError(f"无法安全读取文档包以确认修订状态：{type(exc).__name__}") from exc
    return marks


def _revision_projection(doc) -> tuple:
    """Revision markers present in the body, so no write can hide behind them."""
    counts: dict[str, int] = {}
    for element in doc.element.body.iter():
        name = _local_name(element)
        if name in _REVISION_ALL_MARKS:
            counts[name] = counts.get(name, 0) + 1
    return tuple(sorted(counts.items()))


def _word_state(doc) -> tuple:
    def paras(paragraphs):
        return tuple((p.text, p.style.name, tuple((r.text, r.style.name if r.style else None,
            r.bold, r.italic, r.underline) for r in p.runs)) for p in paragraphs)
    def tables(tbs):
        return tuple(tuple(tuple((c.text, paras(c.paragraphs)) for c in row.cells)
                           for row in t.rows) for t in tbs)
    sections = []
    for section in doc.sections:
        header = _defined_header_footer(section, "header")
        footer = _defined_header_footer(section, "footer")
        sections.append((paras(header.paragraphs) if header is not None else (),
                         tables(header.tables) if header is not None else (),
                         paras(footer.paragraphs) if footer is not None else (),
                         tables(footer.tables) if footer is not None else ()))
    # Revision markers are part of the document's semantics: comparing them makes a write
    # that only removed markers (a formatting revision) verifiable rather than "unchanged".
    return (paras(doc.paragraphs), tables(doc.tables), tuple(sections),
            _revision_projection(doc))


def _has_formula_caches(path: Path) -> bool:
    """openpyxl clears calculated caches on save; reject instead of silently losing them."""
    ns = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
    with zipfile.ZipFile(path) as archive:
        for name in archive.namelist():
            if re.fullmatch(r"xl/worksheets/sheet\d+\.xml", name):
                root = ET.fromstring(archive.read(name))
                for cell in root.iter(ns + "c"):
                    formula = cell.find(ns + "f")
                    cache = cell.find(ns + "v")
                    if formula is not None and cache is not None and cache.text is not None:
                        return True
    return False


def _trimmed_excel_value(cell, request: ExcelTrim) -> str:
    if cell.data_type == "f" or not isinstance(cell.value, str):
        raise GuardError(_TRIM_REMEDY.format(
            reason=f"{cell.coordinate} is a formula or a non-text value, and formulas are never trimmed"))
    if cell.value != request.expected_original:
        raise GuardError(_TRIM_REMEDY.format(
            reason=f"expected_original {request.expected_original!r} does not match the current "
                   f"value of {cell.coordinate}: {cell.value!r}"))
    trimmed = cell.value.strip()
    if trimmed.startswith("="):
        raise GuardError(_TRIM_REMEDY.format(
            reason=f"trimming {cell.coordinate} would produce a formula"))
    if trimmed == cell.value:
        raise GuardError(_TRIM_REMEDY.format(
            reason=f"{cell.coordinate} is already clean, so this edit would change nothing"))
    return trimmed


def _range_trim_values(ws, request: ExcelRangeTrim) -> tuple[dict[str, str], list[str], int]:
    """Cells that whitespace cleaning would change, plus what it deliberately skipped.

    Formulas, numbers and already-clean text are never touched; they are reported as
    skipped so an empty change list can never be mistaken for "nothing was examined".
    """
    changed: dict[str, str] = {}
    skipped: list[str] = []
    examined = 0
    for row in ws[f"{request.first_cell}:{request.last_cell}"]:
        for cell in row:
            examined += 1
            if cell.__class__.__name__ == "MergedCell":
                skipped.append(cell.coordinate)
                continue
            if cell.data_type == "f" or cell.value is None:
                skipped.append(cell.coordinate)
                continue
            if not isinstance(cell.value, str):
                skipped.append(cell.coordinate)
                continue
            trimmed = cell.value.strip()
            if trimmed == cell.value:
                skipped.append(cell.coordinate)
                continue
            if trimmed.startswith("="):
                raise GuardError(f"{cell.coordinate} would become a formula when trimmed")
            changed[cell.coordinate] = trimmed
    return changed, skipped, examined


def _range_set_values(ws, request: ExcelRangeSet) -> tuple[dict[str, Any], dict[str, Any]]:
    """Cells that would change plus their current values, or a whole-edit refusal.

    Every target cell is checked before any change is built, which is what makes the edit
    atomic: one mismatched original refuses the entire block instead of applying the rest.
    """
    left, top, _, _ = range_boundaries(f"{request.first_cell}:{request.last_cell}")
    changed: dict[str, Any] = {}
    before: dict[str, Any] = {}
    for row_offset, (expected_row, value_row) in enumerate(zip(request.expected, request.values)):
        for column_offset, (expected, value) in enumerate(zip(expected_row, value_row)):
            cell = ws.cell(top + row_offset, left + column_offset)
            coordinate = cell.coordinate
            if cell.__class__.__name__ == "MergedCell" or coordinate in ws.merged_cells:
                raise GuardError(_RANGE_SET_REMEDY.format(
                    reason=f"{coordinate} belongs to a merged range, and merged cells are never block targets"))
            if cell.data_type == "f":
                raise GuardError(_RANGE_SET_REMEDY.format(
                    reason=f"{coordinate} holds the formula {cell.value!r}, and a block write never overwrites "
                           f"a formula"))
            if cell.value != expected or type(cell.value) is not type(expected):
                raise GuardError(_RANGE_SET_REMEDY.format(
                    reason=f"{coordinate} is {cell.value!r}, but expected {expected!r}"))
            if type(value) is not type(cell.value) or value != cell.value:
                changed[coordinate] = value
                before[coordinate] = cell.value
    if not changed:
        raise GuardError("block write would make no change: every declared value already matches the file")
    return changed, before


def _sort_values(ws, request: ExcelSort) -> tuple[dict[str, Any], dict[str, Any]]:
    """Compute a bounded value permutation; never mutate the source worksheet.

    Formats remain at coordinates, so each sorted column must have uniform cell
    formatting and every record must have the same row presentation. Blank keys
    sort last in either direction; ties keep their original order.
    """
    left, top, right, bottom = range_boundaries(f"{request.first_cell}:{request.last_cell}")
    key_column, _, _, _ = range_boundaries(request.key_cell)
    column_styles = {}
    row_presentations = set()
    for row in ws.iter_rows(min_row=top, max_row=bottom, min_col=left, max_col=right):
        dimension = ws.row_dimensions.get(row[0].row)
        row_presentations.add((dimension.height, bool(dimension.hidden), dimension.outlineLevel,
                               bool(dimension.collapsed), dimension.style_id)
                              if dimension is not None else (None, False, 0, False, 0))
        for cell in row:
            if cell.coordinate in ws.merged_cells or cell.data_type in {"f", "e"}:
                raise GuardError(f"sort refuses {cell.coordinate}: formula, error or merged cell; "
                                 "select an unmerged literal-only range")
            if cell.comment is not None or cell.hyperlink is not None:
                raise GuardError(f"sort refuses {cell.coordinate}: attached comments/hyperlinks require "
                                 "an Office adapter; use an ordinary literal-only copy")
            if type(cell.value) not in (str, int, float, bool, type(None)):
                raise GuardError(f"sort refuses {cell.coordinate}: unsupported literal type; "
                                 "convert dates or other structured values explicitly first")
            if isinstance(cell.value, float) and not math.isfinite(cell.value):
                raise GuardError(f"sort refuses {cell.coordinate}: non-finite numeric value")
            previous = column_styles.setdefault(cell.column, cell.style_id)
            if previous != cell.style_id:
                raise GuardError(f"sort refuses {cell.coordinate}: column formatting differs between records; "
                                 "use uniformly formatted literal rows or an Office adapter")
    if len(row_presentations) != 1:
        raise GuardError("sort refuses differing row heights/hidden/grouping; use uniformly presented rows")
    key_offset = key_column - left
    records = request.expected
    populated = [row for row in records if row[key_offset] is not None]
    blank = [row for row in records if row[key_offset] is None]
    kinds = {"number" if type(row[key_offset]) in (int, float) else type(row[key_offset]).__name__
             for row in populated}
    if len(kinds) > 1:
        raise GuardError(f"sort refuses {request.key_cell}: mixed key types are ambiguous; "
                         "use one numeric, text or boolean key type")
    values = sorted(populated, key=lambda row: row[key_offset], reverse=request.descending) + blank
    block = ExcelRangeSet(path=request.path, sheet=request.sheet,
                          first_cell=request.first_cell, last_cell=request.last_cell,
                          expected=request.expected, values=values)
    return _range_set_values(ws, block)


def _written_value_matches(state_entry, value: Any) -> bool:
    """Compare a reopened cell against the intended literal without type coercion."""
    if state_entry is None:
        return value is None
    return type(state_entry[0]) is type(value) and state_entry[0] == value


def _excel_state(wb) -> dict:
    return {ws.title: {
        "cells": {c.coordinate: (c.value, c.data_type, c.style_id) for c in ws._cells.values()
                  if c.value is not None or c.has_style},
        "merged": tuple(str(r) for r in ws.merged_cells.ranges),
        "charts": len(ws._charts), "images": len(ws._images),
    } for ws in wb.worksheets}


def _replace_runs(paragraph, old: str, new: str) -> int:
    runs = paragraph.runs
    chars = "".join(r.text for r in runs)
    matches = []
    cursor = 0
    while (at := chars.find(old, cursor)) >= 0:
        matches.append((at, at + len(old)))
        cursor = at + len(old)
    # Apply backwards so previously located run offsets are still valid.
    positions = [(i, j) for i, run in enumerate(runs) for j in range(len(run.text))]
    touched = {positions[pos][0] for start, end in matches for pos in range(start, end)}
    for index in touched:
        if any(child.tag not in {qn("w:rPr"), qn("w:t")} for child in runs[index]._r):
            raise GuardError("target run contains non-text XML; preserving anchors/tabs requires Office adapter")
    for start, end in reversed(matches):
        first_i, first_j = positions[start]
        last_i, last_j = positions[end - 1]
        if first_i == last_i:
            text = runs[first_i].text
            runs[first_i].text = text[:first_j] + new + text[last_j + 1:]
        else:
            first = runs[first_i].text
            last = runs[last_i].text
            runs[first_i].text = first[:first_j] + new
            for i in range(first_i + 1, last_i):
                runs[i].text = ""
            runs[last_i].text = last[last_j + 1:]
    return len(matches)


def _plain_paragraph(paragraph) -> None:
    """Do not silently skip hyperlinks, tracked edits, fields, or other XML."""
    if any(child.tag not in {qn("w:pPr"), qn("w:r")} for child in paragraph._p):
        raise GuardError("target paragraph contains hyperlink/unsupported XML")
    if any(any(child.tag not in {qn("w:rPr"), qn("w:t")} for child in run._r)
           for run in paragraph.runs):
        raise GuardError("target paragraph contains non-text run XML")


def _table_paragraph(doc, request: WordTableCellReplace):
    if request.table_index >= len(doc.tables):
        raise GuardError("table index out of bounds")
    table = doc.tables[request.table_index]
    if request.row_index >= len(table.rows) or request.cell_index >= len(table.columns):
        raise GuardError("table cell index out of bounds")
    cell = table.cell(request.row_index, request.cell_index)
    tc_pr = cell._tc.tcPr
    if tc_pr is not None and (tc_pr.find(qn("w:gridSpan")) is not None or
                              tc_pr.find(qn("w:vMerge")) is not None):
        raise GuardError("merged table cells require a high-fidelity adapter")
    if request.paragraph_index >= len(cell.paragraphs):
        raise GuardError("cell paragraph index out of bounds")
    paragraph = cell.paragraphs[request.paragraph_index]
    _plain_paragraph(paragraph)
    return paragraph


def _format_target(doc, request: WordFormat):
    if request.paragraph_index >= len(doc.paragraphs):
        raise GuardError("paragraph index out of bounds")
    paragraph = doc.paragraphs[request.paragraph_index]
    _plain_paragraph(paragraph)
    if request.run_index >= len(paragraph.runs):
        raise GuardError("run index out of bounds")
    run = paragraph.runs[request.run_index]
    if not run.text or len(run._r.findall(qn("w:t"))) != 1:
        raise GuardError("formatting requires one existing plain-text run")
    return run


def _format_values(run) -> str:
    return repr({"text": run.text, "bold": run.bold,
                 "italic": run.italic, "underline": run.underline})


def _apply_format(run, request: WordFormat) -> None:
    for property_name in ("bold", "italic", "underline"):
        value = getattr(request, property_name)
        if value is not None:
            setattr(run, property_name, value)


def _task_literal_audit(audit):
    verification = audit.get('task_route_verification')
    return (audit.get('kind') in {'word_batch_replace', 'word_table_cell_replace'} and
            type(verification) is dict and verification.get('status') == 'passed')


@dataclass(frozen=True)
class Plan:
    id: str
    request: Edit
    source_sha256: str
    target: str
    before: str
    after: str
    risks: tuple[str, ...]
    confirmation_required: bool
    task_validation: Any | None = None

    def public(self) -> dict:
        return {"id": self.id, "kind": self.request.kind, "source": self.request.path,
                "source_sha256": self.source_sha256, "target": self.target,
                "before": self.before, "after": self.after, "risks": list(self.risks),
                "confirmation_required": self.confirmation_required}


class Editor:
    def __init__(self, authorized_root: Path, *, legacy_backend: Any | None = None,
                 require_confirmation: bool = False, render_validation: bool = False):
        """``legacy_backend`` is the injected Office COM channel for .doc/.xls edits.

        It stays ``None`` unless the caller supplies one, and every legacy plan then refuses
        with the reason the backend's capability probe reports: installation alone must
        never open a write channel.
        """
        self.space = Workspace(authorized_root)
        self.plans: dict[str, Plan] = {}
        self.audit = self.space.root / ".office_agent"
        self.legacy_backend = legacy_backend
        if type(require_confirmation) is not bool:
            raise GuardError("confirmation policy must be boolean")
        self.require_confirmation = require_confirmation
        if type(render_validation) is not bool:
            raise GuardError('render validation policy must be boolean')
        self.render_validation = render_validation
        self._legacy_probes: dict[str, dict[str, Any]] = {}

    def audit_summary(self) -> dict[str, Any]:
        """Human-facing, read-only audit inventory; never registered as a model tool."""
        return audit_summary(self.space.root)

    def retain_audits(self, *, older_than_days: int = 90, apply: bool = False) -> dict[str, Any]:
        """Explicit human maintenance entry; preview only unless apply is exactly True."""
        return retain_audits(self.space.root, older_than_days=older_than_days, apply=apply)

    def _legacy_channel(self, request: Edit) -> Any:
        """The Office backend for one legacy request, or a refusal that states why."""
        application = "word" if isinstance(request, LegacyWordReplace) else "excel"
        if self.legacy_backend is None:
            raise GuardError(
                "本会话未启用 .doc/.xls 编辑通道（未注入 Office COM 后端）；源文件未被改动。"
                "普通 docx/xlsx 请使用 Python 通道")
        probe = self._legacy_probes.get(application)
        if probe is None:
            # Probed once per Editor: it starts an Office instance and creates a blank
            # document of our own, saves a throwaway copy and reopens it.
            try:
                probe = self.legacy_backend.probe(application)
            except Exception as exc:
                probe = {"available": False, "stage": "startup",
                         "reason": f"{type(exc).__name__}: {str(exc)[:200]}"}
            if not isinstance(probe, dict):
                probe = {"available": False, "stage": "startup",
                         "reason": "Office 后端返回了不可用的探测结果"}
            self._legacy_probes[application] = probe
        if not probe.get("available"):
            from .com import describe_write_refusal
            raise GuardError(describe_write_refusal(application, probe))
        return self.legacy_backend

    def _legacy_working_copy(self, candidate: Path) -> Path:
        """A legacy edit may only ever target the audit job's own working copy.

        This is the machine-checked form of "never edit the source in place": even a bug in
        the caller cannot make the COM channel write to the user's file.
        """
        audit = self.audit.resolve(strict=False)
        resolved = candidate.resolve(strict=False)
        if not (resolved.is_relative_to(audit) and resolved != audit):
            raise GuardError("legacy edits may only target the audit job's own copy")
        return candidate

    def plan(self, request: Edit) -> Plan:
        from .advanced_writes import REQUESTS, make_plan
        if isinstance(request, REQUESTS):
            return make_plan(self, request)
        src = self.space.file(request.path)
        excel_kinds = (ExcelSet, ExcelTrim, ExcelRangeTrim, ExcelRangeSet, ExcelSort, ExcelFormulaFill, ExcelBarChart, ExcelChart)
        legacy = isinstance(request, _LEGACY_KINDS)
        if legacy:
            required = ".doc" if isinstance(request, LegacyWordReplace) else ".xls"
            if src.suffix.lower() != required:
                raise GuardError(f"{required} required for this legacy edit")
            # Legacy OLE files are unreadable to the OOXML inspector by design; the COM
            # channel has its own preconditions instead of that risk gate.
            _require_legacy_source(src)
        elif src.suffix.lower() != (".xlsx" if isinstance(request, excel_kinds) else ".docx"): 
            raise GuardError("only ordinary docx/xlsx edits; legacy, macro and encrypted files are read-only")
        # The query inspector must report unknown structures; writes are conservative.
        from .documents import classify_document, inspect_file
        if not legacy:
            report = inspect_file(src)
            unsafe = list(report.get("risks", []))
            if isinstance(request, WordRevisionDecision):
                # Deciding revisions is the one operation whose target *is* the advanced XML the
                # inspector flags. That is tolerated only while w:ins/w:del are the sole suspicious
                # features: any other advanced structure keeps this edit refused as before.
                features = set(classify_document(src).suspicious_features)
                if features and features <= _REVISION_TOLERATED_FEATURES:
                    unsafe = [risk for risk in unsafe
                              if risk not in _REVISION_TOLERATED_FEATURES
                              and risk != "suspicious_package_features_present"]
            if unsafe or report.get("protected") or report.get("macros") or report.get("external_links"):
                raise GuardError(f"unsupported/risky structures require a high-fidelity adapter: {unsafe}")
        if isinstance(request, WordReplace):
            if request.old == request.new:
                raise GuardError("replacement would make no change")
            doc = Document(src)
            if request.paragraph_index >= len(doc.paragraphs):
                raise GuardError("paragraph index out of bounds")
            before = doc.paragraphs[request.paragraph_index].text
            count = before.count(request.old)
            if count != request.expected_count:
                raise GuardError(f"expected {request.expected_count} occurrences, found {count}")
            paragraph = doc.paragraphs[request.paragraph_index]
            if _replace_runs(paragraph, request.old, request.new) != count:
                raise GuardError("target text includes hyperlink/unsupported run content")
            after = before.replace(request.old, request.new)
            if paragraph.text != after:
                raise GuardError("paragraph text cannot be replaced without changing unsupported structures")
            target = f"paragraph[{request.paragraph_index}]"
        elif isinstance(request, WordBatchReplace):
            if request.old == request.new:
                raise GuardError("replacement would make no change")
            doc = Document(src)
            targets = _batch_replace(doc, request)
            before = json.dumps([{"paragraph_index": t["paragraph_index"], "before": t["before"]}
                                 for t in targets], ensure_ascii=False)
            after = json.dumps([{"paragraph_index": t["paragraph_index"], "after": t["after"]}
                                for t in targets], ensure_ascii=False)
            target = (f"{sum(t['before'].count(request.old) for t in targets)} occurrence(s) in "
                      f"{len(targets)} paragraph(s), capped at {request.max_occurrences}")
        elif isinstance(request, WordTableRow):
            doc = Document(src)
            _simple_table_row(doc, request)  # full bounds/merge validation before previewing
            before = repr(_table_row_summary(doc, request, describe_deleted=True))
            _apply_table_row(doc, request)
            after = repr(_table_row_summary(doc, request))
            target = f"table[{request.table_index}].row[{request.row_index}] {request.action}"
            if before == after:
                raise GuardError("table row operation would make no change")
        elif isinstance(request, WordFormat):
            doc = Document(src)
            run = _format_target(doc, request)
            before = _format_values(run)
            _apply_format(run, request)
            after = _format_values(run)
            if before == after:
                raise GuardError("formatting would make no change")
            target = f"paragraph[{request.paragraph_index}].run[{request.run_index}]"
        elif isinstance(request, WordTableCellReplace):
            doc = Document(src)
            paragraph = _table_paragraph(doc, request)
            before = paragraph.text
            if before.count(request.old) != 1:
                raise GuardError("expected exactly one occurrence in target cell paragraph")
            if _replace_runs(paragraph, request.old, request.new) != 1:
                raise GuardError("table cell replacement count changed")
            after = before.replace(request.old, request.new)
            if paragraph.text != after or before == after:
                raise GuardError("table cell cannot be replaced without changing unsupported structures")
            target = (f"table[{request.table_index}].row[{request.row_index}]"
                      f".cell[{request.cell_index}].paragraph[{request.paragraph_index}]")
        elif isinstance(request, DocxMerge):
            sources = _merge_sources(self.space, src, request.sources)
            if src not in sources:
                raise GuardError("the authorized file must be one of the merge sources; "
                                 "list it by name in sources")
            before, after, totals = _merge_plan_value(sources)
            target = (f"按顺序合并 {len(sources)} 个 DOCX（{', '.join(request.sources)}）→ 新输出副本 "
                      f"{src.stem}.agent-<计划ID>{src.suffix}；所有源文件保持原样"
                      + (f"；含 {totals['numbered']} 个有效编号段落（含样式继承），编号定义与样式将迁移并隔离重映射"
                         if totals["numbered"] else ""))
        elif isinstance(request, DocxSplit):
            before, after, titles, totals = _split_plan_value(src, request.level)
            if len(titles) <= 1:
                raise GuardError(f"按 Heading ≤{request.level} 拆分只会得到 1 个部分，无需拆分；"
                                 f"请确认标题使用了 Heading 样式")
            if len(titles) > _MERGE_MAX_PARTS:
                raise GuardError(f"拆分会产生 {len(titles)} 个输出，超过一次确认的上限 "
                                 f"{_MERGE_MAX_PARTS}")
            target = (f"按 Heading ≤{request.level} 拆分为 {len(titles)} 个新 DOCX："
                      f"{'、'.join(titles[:5])}{' 等' if len(titles) > 5 else ''}；源文件保持原样")
        elif isinstance(request, LegacyWordReplace):
            backend = self._legacy_channel(request)
            paragraphs = _legacy_paragraph_window(backend, src)
            if request.paragraph_index >= len(paragraphs):
                raise GuardError(f"paragraph index {request.paragraph_index} is out of bounds for the "
                                 f"{len(paragraphs)} paragraph(s) read (bounded window "
                                 f"{_MAX_LEGACY_PARAGRAPH_WINDOW})")
            before = paragraphs[request.paragraph_index]
            count = before.count(request.old)
            if count != request.expected_count:
                raise GuardError(f"expected {request.expected_count} occurrence(s) of the literal, found {count}")
            after = before.replace(request.old, request.new)
            if before == after:
                raise GuardError("replacement would make no change")
            target = (f"paragraph[{request.paragraph_index}] (.doc 经 Office COM；"
                      f"已核对窗口内 {len(paragraphs)} 段)")
        elif isinstance(request, LegacyCellSet):
            backend = self._legacy_channel(request)
            current = backend.read_excel_cell(str(src), sheet=request.sheet, cell=request.cell)
            if not _legacy_same_value(current, request.expected_original):
                raise GuardError(f"expected_original {request.expected_original!r} does not match "
                                 f"{request.sheet}!{request.cell}: {current!r}")
            if _legacy_same_value(current, request.value):
                raise GuardError("cell assignment would make no change")
            before = repr(current)
            after = repr(request.value)
            target = f"{request.sheet}!{request.cell} (.xls 经 Office COM)"
        elif isinstance(request, WordMergedCellText):
            doc = Document(src)
            table, cell, info = _merged_cell(doc, request)
            merges_before = _merge_projection(doc)
            # The preview names the exact merge it writes into, so the user can check it.
            merge_label = ((f" gridSpan={info['grid_span']}" if info["grid_span"] > 1 else "")
                           + (f" vMerge={info['vmerge']}" if info["vmerge"] else ""))
            if request.action == "replace":
                before, after = _merged_cell_replace(cell, request)
                target = (f"table[{request.table_index}] 合并单元格 逻辑(行 {request.row_index}, "
                          f"列 {request.column_index}) 段落[{request.paragraph_index}]{merge_label}")
            else:
                before, after = _merged_cell_append(cell, request)
                target = (f"table[{request.table_index}] 合并单元格 逻辑(行 {request.row_index}, "
                          f"列 {request.column_index}) 末尾追加段落{merge_label}")
            if _merge_projection(doc) != merges_before:
                raise GuardError("编辑会改变表格的合并结构，已拒绝")
        elif isinstance(request, WordRevisionDecision):
            doc = Document(src)
            inventory = _revision_inventory(doc.element.body)
            unsupported = {name: count for name, count in inventory["unsupported"].items() if count}
            if unsupported:
                raise GuardError(_REVISION_REMEDY.format(reason="文档含有本操作不支持的修订标记：" +
                    "、".join(f"{name}×{count}" for name, count in sorted(unsupported.items()))))
            outside = _package_revision_marks(src, include_document=False)
            if outside:
                raise GuardError(_REVISION_REMEDY.format(
                    reason="正文之外的部件中仍有修订标记（" +
                           "、".join(f"{name}×{count}" for name, count in sorted(outside.items())) +
                           "），本操作只处理正文 word/document.xml"))
            total = (len(inventory["insertions"]) + len(inventory["deletions"])
                     + len(inventory["format_changes"]))
            if not total:
                raise GuardError("文档正文中没有可处理的修订标记（w:ins/w:del/rPrChange/pPrChange）")
            measured = _revision_apply_and_measure(doc, request.action)
            before, after = measured["before"], measured["after"]
            target = (f"{'接受' if request.action == 'accept_all' else '拒绝'}全部修订："
                      f"插入 {len(inventory['insertions'])} 处、删除 {len(inventory['deletions'])} 处、"
                      f"格式修订 {len(inventory['format_changes'])} 处；只处理正文 document.xml")
        else:
            if isinstance(request, (ExcelBarChart, ExcelChart, ExcelFormulaFill, ExcelSort)):
                operation = ("literal row sorting" if isinstance(request, ExcelSort) else
                             "formula fill" if isinstance(request, ExcelFormulaFill) else "chart creation")
                _chart_package_safe(src, operation)
            if _has_formula_caches(src):
                raise GuardError("formula cached results would be lost by openpyxl; use a verified Office adapter")
            wb = load_workbook(src, data_only=False)
            try:
                if request.sheet not in wb:
                    raise GuardError("sheet not found")
                if isinstance(request, (ExcelBarChart, ExcelChart)):
                    if len(wb.worksheets) != 1 or wb.active._charts or wb.active._images:
                        raise GuardError("chart creation requires one sheet and no existing charts/drawings")
                    references, categories = _chart_sources(wb[request.sheet], request)
                    target = f"{request.sheet}!{request.anchor}"
                    before = "no chart"
                    after = repr({"type": _CHART_TYPES[getattr(request, "chart_type", "bar")],
                                  "title": request.title,
                                  "series": [_reference_formula(reference) for reference in references],
                                  "categories": request.categories_range,
                                  "axis_titles": (getattr(request, "category_axis_title", None),
                                                  getattr(request, "value_axis_title", None)),
                                  "legend_position": getattr(request, "legend_position", "right"),
                                  "anchor": request.anchor})
                    overwrites_formula = False
                    cell = None
                elif isinstance(request, ExcelRangeTrim):
                    changed, skipped, examined = _range_trim_values(wb[request.sheet], request)
                    if not changed:
                        raise GuardError(f"no cell in {request.first_cell}:{request.last_cell} needs "
                                         f"whitespace cleaning ({examined} cells examined)")
                    target = f"{request.sheet}!{request.first_cell}:{request.last_cell}"
                    before = repr({coord: wb[request.sheet][coord].value for coord in changed})
                    after = repr(changed)
                    overwrites_formula = False
                    cell = None
                elif isinstance(request, (ExcelRangeSet, ExcelSort)):
                    changed, before_values = (_sort_values(wb[request.sheet], request) if isinstance(request, ExcelSort)
                                              else _range_set_values(wb[request.sheet], request))
                    target = f"{request.sheet}!{request.first_cell}:{request.last_cell}"
                    before = repr(before_values)
                    after = repr(changed)
                    overwrites_formula = False
                    cell = None
                elif isinstance(request, ExcelFormulaFill):
                    values = _formula_fill_values(wb[request.sheet], request)
                    target = f"{request.sheet}!{request.target_range}"
                    before = repr({coord: None for coord in values})
                    after = repr(values)
                    overwrites_formula = False
                    cell = None
                else:
                    cell = wb[request.sheet][request.cell]
                if cell is not None:
                    if cell.coordinate in wb[request.sheet].merged_cells and cell.__class__.__name__ == "MergedCell":
                        raise GuardError("cannot edit non-anchor merged cell")
                    if isinstance(request, ExcelTrim):
                        value = _trimmed_excel_value(cell, request)
                    else:
                        value = request.value
                        if cell.value == value and type(cell.value) is type(value):
                            raise GuardError("cell assignment would make no change")
                    before = repr(cell.value)
                    after = repr(value)
                    target = f"{request.sheet}!{cell.coordinate}"
                    overwrites_formula = cell.data_type == "f"
            finally:
                wb.close()
        risks = (("chart creation requires human confirmation",) if isinstance(request, (ExcelBarChart, ExcelChart)) else
                 ("formula fill requires human confirmation; results are not calculated",) if isinstance(request, ExcelFormulaFill) else
                  ("literal row sorting requires its own human confirmation; exclude headers and include all record columns; "
                   "formatting stays at coordinates; blank keys sort last and ties are stable",) if isinstance(request, ExcelSort) else
                  ("block write of literal cells requires its own human confirmation",) if isinstance(request, ExcelRangeSet) else
                  ("bulk whitespace cleaning requires its own human confirmation",) if isinstance(request, ExcelRangeTrim) else
                  ("merged-cell text edit requires its own human confirmation",) if isinstance(request, WordMergedCellText) else
                  ("merge/split create new documents and require their own human confirmation; "
                   "no source file is modified", "B7: styles/numbering migrated with isolated IDs; "
                    "headers/footers/footnotes/endnotes/comments refused if present; accepted inputs absent; "
                    "structural verification only, visual fidelity unverified") if isinstance(request, _MERGE_KINDS) else
                  ("legacy .doc/.xls edit through Office COM requires its own human confirmation",) if isinstance(request, _LEGACY_KINDS) else
                  ("deciding all revisions requires its own human confirmation",) if isinstance(request, WordRevisionDecision) else
                  ("deletes a table row and requires its own human confirmation",) if isinstance(request, WordTableRow) and request.action == "delete" else
                  ("batch replacement requires its own human confirmation",) if isinstance(request, WordBatchReplace) else
                  ("batch replacement",) if isinstance(request, WordReplace) and count > 1 else
                 ("overwrites existing formula",) if isinstance(request, ExcelSet) and overwrites_formula else ())
        plan = Plan(uuid.uuid4().hex, request, digest(src), target, before, after,
                    risks, bool(risks) or self.require_confirmation)
        self.plans[plan.id] = plan
        return plan

    def execute(self, plan_id: str, *, confirmed: bool = False) -> dict:
        from .advanced_writes import REQUESTS, execute
        if isinstance(self.plans[plan_id].request, REQUESTS):
            return execute(self, self.plans[plan_id], confirmed)
        # confirmed must originate from a separate human UI action, never from model tool arguments.
        plan = self.plans.pop(plan_id)
        if (self.require_confirmation or plan.confirmation_required) and confirmed is not True:
            self.plans[plan_id] = plan
            raise GuardError("explicit human confirmation required")
        src = self.space.file(plan.request.path)
        if digest(src) != plan.source_sha256:
            raise GuardError("source changed since preview; create a new plan")
        self.audit.mkdir(exist_ok=True)
        if self.audit.is_symlink() or not self.audit.resolve(strict=True).is_relative_to(self.space.root):
            raise GuardError("audit directory redirected outside authorized root")
        # Detect a pre-existing output before creating the per-job audit directory;
        # a collision must not leave an empty task with no backup or audit record.
        output = self.space.file(str(src.with_name(f"{src.stem}.agent-{plan.id[:8]}{src.suffix}")), existing=False)
        job = self.audit / plan.id
        job.mkdir(exist_ok=False)
        backup = job / ("backup" + src.suffix)
        temp = job / ("working" + src.suffix)
        log = {**plan.public(), "status": "failed", "backup": str(backup), "output": str(output)}
        published = False
        audited = False
        published_paths: list[Path] = []
        creation_plan: list[tuple[Path, Path]] | None = None
        try:
            shutil.copy2(src, backup)
            if digest(backup) != plan.source_sha256:
                raise GuardError("backup verification failed")
            if isinstance(plan.request, _MERGE_KINDS):
                # Merge/split only ever create new documents: no source is written, and the
                # published outputs are new files whose sources stay byte-identical.
                creation_plan, creation_info = self._creation_outputs(plan, src, job, temp)
                log.update(creation_info)
            elif isinstance(plan.request, _LEGACY_KINDS):
                # Legacy .doc/.xls: the COM channel edits only the audit job's own working
                # copy, is verified with the same reopened-file check, and never touches src.
                backend = self._legacy_channel(plan.request)
                working = self._legacy_working_copy(temp)
                shutil.copy2(backup, working)
                if isinstance(plan.request, LegacyWordReplace):
                    between = _legacy_paragraph_window(backend, backup)
                    if (plan.request.paragraph_index >= len(between)
                            or between[plan.request.paragraph_index] != plan.before):
                        raise GuardError("legacy paragraph target changed since preview")
                    written = backend.replace_word_paragraph(
                        str(working), index=plan.request.paragraph_index, old=plan.request.old,
                        new=plan.request.new, expected_count=plan.request.expected_count)
                    if written != plan.after:
                        raise GuardError("legacy paragraph replacement does not match the preview")
                    reopened = _legacy_paragraph_window(backend, working)
                    if len(reopened) != len(between):
                        raise GuardError("legacy 段落数量发生变化")
                    for index, (was, now) in enumerate(zip(between, reopened)):
                        if index != plan.request.paragraph_index and was != now:
                            raise GuardError(f"legacy 非目标段落发生变化：paragraph[{index}]")
                    log["legacy_verification"] = {
                        "kind": "paragraph_window", "window": len(between),
                        "target": plan.request.paragraph_index}
                else:
                    current = backend.read_excel_cell(str(backup), sheet=plan.request.sheet,
                                                      cell=plan.request.cell)
                    if not _legacy_same_value(current, plan.request.expected_original):
                        raise GuardError("legacy cell target changed since preview")
                    before_window = _legacy_region_values(backend, backup, plan.request.sheet)
                    written_value = backend.set_excel_cell(str(working), sheet=plan.request.sheet,
                                                           cell=plan.request.cell, value=plan.request.value)
                    if not _legacy_same_value(written_value, plan.request.value):
                        raise GuardError("legacy cell assignment does not match the preview")
                    after_window = _legacy_region_values(backend, working, plan.request.sheet)
                    _legacy_verify_window(before_window, after_window,
                                          excluded=plan.request.cell, label="单元格")
                    log["legacy_verification"] = {"kind": "cell_window", "window": _LEGACY_REGION,
                                                  "target": f"{plan.request.sheet}!{plan.request.cell}"}
            elif not isinstance(plan.request, (ExcelSet, ExcelTrim, ExcelRangeTrim, ExcelRangeSet, ExcelSort,
                                               ExcelFormulaFill, ExcelBarChart, ExcelChart)):
                doc = Document(backup)
                before = _word_state(doc)
                merges_before = None
                revision_changes = None
                if isinstance(plan.request, WordReplace):
                    p = doc.paragraphs[plan.request.paragraph_index]
                    if _replace_runs(p, plan.request.old, plan.request.new) != plan.request.expected_count:
                        raise GuardError("replacement count changed")
                    preview_after = p.text
                elif isinstance(plan.request, WordBatchReplace):
                    targets = _batch_replace(doc, plan.request)
                    preview_after = json.dumps([{"paragraph_index": t["paragraph_index"],
                                                 "after": t["after"]} for t in targets],
                                               ensure_ascii=False)
                elif isinstance(plan.request, WordTableRow):
                    if repr(_table_row_summary(doc, plan.request, describe_deleted=True)) != plan.before:
                        raise GuardError("table row target changed since preview")
                    _apply_table_row(doc, plan.request)
                    if repr(_table_row_summary(doc, plan.request)) != plan.after:
                        raise GuardError("table row operation does not match the preview")
                    preview_after = plan.after
                elif isinstance(plan.request, WordFormat):
                    run = _format_target(doc, plan.request)
                    if _format_values(run) != plan.before:
                        raise GuardError("format target changed since preview")
                    _apply_format(run, plan.request)
                    preview_after = _format_values(run)
                elif isinstance(plan.request, WordMergedCellText):
                    merges_before = _merge_projection(doc)
                    _, cell, _ = _merged_cell(doc, plan.request)
                    if plan.request.action == "replace":
                        paragraph = _merged_cell_paragraph(cell, plan.request)
                        if paragraph.text != plan.before:
                            raise GuardError("merged cell target changed since preview")
                        if _replace_runs(paragraph, plan.request.old,
                                         plan.request.new) != plan.request.expected_count:
                            raise GuardError("merged cell replacement count changed")
                        preview_after = paragraph.text
                    else:
                        _, preview_after = _merged_cell_append(cell, plan.request)
                    if _merge_projection(doc) != merges_before:
                        raise GuardError("合并结构在写入过程中被改变")
                elif isinstance(plan.request, WordRevisionDecision):
                    measured = _revision_apply_and_measure(doc, plan.request.action)
                    if measured["before"] != plan.before:
                        raise GuardError("revision state changed since preview")
                    if measured["after"] != plan.after:
                        raise GuardError("revision decision does not match the preview")
                    revision_changes = measured
                    preview_after = measured["after"]
                else:
                    p = _table_paragraph(doc, plan.request)
                    if p.text != plan.before or _replace_runs(p, plan.request.old, plan.request.new) != 1:
                        raise GuardError("table cell target changed since preview")
                    preview_after = p.text
                if preview_after != plan.after:
                    raise GuardError("Word target mismatch")
                intended = _word_state(doc)
                doc.save(temp)
                saved = Document(temp)
                actual = _word_state(saved)
                if actual != intended or before == actual or digest(src) != plan.source_sha256:
                    raise GuardError("Word post-write verification failed")
                # A merged-cell write must leave every gridSpan/vMerge exactly as it was,
                # and a revision decision must leave no revision mark behind anywhere.
                if merges_before is not None and _merge_projection(saved) != merges_before:
                    raise GuardError("合并结构在写入后与写入前不一致")
                if revision_changes is not None:
                    remaining = _package_revision_marks(temp)
                    if remaining:
                        raise GuardError("写入后文档仍残留修订标记：" +
                                         "、".join(f"{name}×{count}" for name, count in sorted(remaining.items())))
                if isinstance(plan.request, WordReplace):
                    saved_after = saved.paragraphs[plan.request.paragraph_index].text
                elif isinstance(plan.request, WordBatchReplace):
                    # Replacement never adds or removes paragraphs, so the previewed indices
                    # still address the same paragraphs after saving.
                    saved_after = json.dumps([{"paragraph_index": t["paragraph_index"],
                                               "after": saved.paragraphs[t["paragraph_index"]].text}
                                              for t in targets], ensure_ascii=False)
                elif isinstance(plan.request, WordTableRow):
                    saved_after = repr(_table_row_summary(saved, plan.request))
                elif isinstance(plan.request, WordFormat):
                    saved_after = _format_values(_format_target(saved, plan.request))
                elif isinstance(plan.request, WordMergedCellText):
                    _, saved_cell, _ = _merged_cell(saved, plan.request)
                    saved_after = (saved_cell.text if plan.request.action == "append"
                                   else _merged_cell_paragraph(saved_cell, plan.request).text)
                elif isinstance(plan.request, WordRevisionDecision):
                    # Rebuilds the preview value from the reopened file, so a text difference
                    # after the decision fails the generic comparison below.
                    saved_after = _revision_value(_word_text_projection(saved),
                                                  revision_changes["changed"],
                                                  revision_changes["changed_truncated"])
                else:
                    saved_after = _table_paragraph(saved, plan.request).text
                if saved_after != plan.after:
                    raise GuardError("Word target mismatch")
            else:
                wb = load_workbook(backup, data_only=False)
                try:
                    before = _excel_state(wb)
                    if isinstance(plan.request, (ExcelBarChart, ExcelChart)):
                        if len(wb.worksheets) != 1 or wb.active._charts or wb.active._images:
                            raise GuardError("chart source changed since preview")
                        ws = wb[plan.request.sheet]
                        references, categories = _chart_sources(ws, plan.request)
                        chart = _build_chart(plan.request, references, categories)
                        ws.add_chart(chart, plan.request.anchor)
                        # Project the chart we just built and require the reopened file to
                        # match it: this verifies the round trip instead of a hand-written
                        # tuple that can drift from openpyxl's internal representation.
                        expected_chart = _normalized_projection(chart)
                        coordinate = None
                    elif isinstance(plan.request, ExcelRangeTrim):
                        ws = wb[plan.request.sheet]
                        changed, _, _ = _range_trim_values(ws, plan.request)
                        if repr({coord: ws[coord].value for coord in changed}) != plan.before or \
                                repr(changed) != plan.after:
                            raise GuardError("whitespace cleaning target changed since preview")
                        for coord, value in changed.items():
                            ws[coord].value = value
                        coordinate = None
                    elif isinstance(plan.request, (ExcelRangeSet, ExcelSort)):
                        ws = wb[plan.request.sheet]
                        changed, before_values = (_sort_values(ws, plan.request) if isinstance(plan.request, ExcelSort)
                                                  else _range_set_values(ws, plan.request))
                        if repr(before_values) != plan.before or repr(changed) != plan.after:
                            raise GuardError("sort target changed since preview" if isinstance(plan.request, ExcelSort)
                                             else "block write target changed since preview")
                        for coordinate, value in changed.items():
                            ws[coordinate].value = value
                        coordinate = None
                    elif isinstance(plan.request, ExcelFormulaFill):
                        ws = wb[plan.request.sheet]
                        values = _formula_fill_values(ws, plan.request)
                        if repr({coord: None for coord in values}) != plan.before or repr(values) != plan.after:
                            raise GuardError("formula fill changed since preview")
                        for coord, formula in values.items():
                            ws[coord].value = formula
                        coordinate = None
                    else:
                        cell = wb[plan.request.sheet][plan.request.cell]
                        if isinstance(plan.request, ExcelTrim):
                            value = _trimmed_excel_value(cell, plan.request)
                            if repr(cell.value) != plan.before or repr(value) != plan.after:
                                raise GuardError("Excel trim target changed since preview")
                        else:
                            value = plan.request.value
                        cell.value = value
                        coordinate = cell.coordinate
                    intended = _excel_state(wb)
                    wb.save(temp)
                finally:
                    wb.close()
                verified = load_workbook(temp, data_only=False)
                try:
                    actual = _excel_state(verified)
                    if isinstance(plan.request, (ExcelBarChart, ExcelChart)):
                        ws = verified[plan.request.sheet]
                        actual_chart = (_normalized_projection(ws._charts[0])
                                        if len(ws._charts) == 1 else None)
                finally:
                    verified.close()
                if actual != intended or digest(src) != plan.source_sha256:
                    raise GuardError("Excel post-write verification failed")
                if isinstance(plan.request, (ExcelBarChart, ExcelChart)):
                    unchanged = {name: {**state, "charts": 0} for name, state in actual.items()}
                    if unchanged != before or actual_chart != expected_chart:
                        raise GuardError("Excel chart or non-target semantics changed")
                elif isinstance(plan.request, ExcelRangeTrim):
                    cells = actual[plan.request.sheet]["cells"]
                    if any(cells.get(coord, (None, None))[0:2] != (value, "s")
                           for coord, value in changed.items()):
                        raise GuardError("whitespace cleaning target mismatch")
                    # Every other cell, sheet and tracked feature must be untouched.
                    unchanged = {name: {**state, "cells": {coord: cell for coord, cell in state["cells"].items()
                                                if name != plan.request.sheet or coord not in changed}}
                                 for name, state in actual.items()}
                    expected = {name: {**state, "cells": {coord: cell for coord, cell in state["cells"].items()
                                               if name != plan.request.sheet or coord not in changed}}
                                for name, state in before.items()}
                    if unchanged != expected:
                        raise GuardError("whitespace cleaning changed non-target semantics")
                elif isinstance(plan.request, (ExcelRangeSet, ExcelSort)):
                    target_cells = actual[plan.request.sheet]["cells"]
                    if any(not _written_value_matches(target_cells.get(coord), value)
                           for coord, value in changed.items()):
                        raise GuardError("block write target mismatch")
                    # Every other cell, sheet and tracked feature must be untouched.
                    unchanged = {name: {**state, "cells": {coord: cell for coord, cell in state["cells"].items()
                                                if name != plan.request.sheet or coord not in changed}}
                                 for name, state in actual.items()}
                    expected = {name: {**state, "cells": {coord: cell for coord, cell in state["cells"].items()
                                               if name != plan.request.sheet or coord not in changed}}
                                for name, state in before.items()}
                    if unchanged != expected:
                        raise GuardError("block write changed non-target semantics")
                elif isinstance(plan.request, ExcelFormulaFill):
                    target_cells = actual[plan.request.sheet]["cells"]
                    if any(target_cells.get(coord, (None, None))[0:2] != (formula, "f")
                           for coord, formula in values.items()):
                        raise GuardError("formula fill target mismatch")
                    # Verify source/non-target cell semantics and tracked sheet features;
                    # arbitrary workbook metadata is not covered by this projection.
                    unchanged = {name: {**state, "cells": {coord: cell for coord, cell in state["cells"].items()
                                                if name != plan.request.sheet or coord not in values}}
                                 for name, state in actual.items()}
                    expected = {name: {**state, "cells": {coord: cell for coord, cell in state["cells"].items()
                                               if name != plan.request.sheet or coord not in values}}
                                for name, state in before.items()}
                    if unchanged != expected:
                        raise GuardError("formula fill changed non-target semantics")
                elif actual[plan.request.sheet]["cells"].get(coordinate, (None,))[0] != value:
                    raise GuardError("Excel target mismatch")
            if plan.task_validation is not None:
                from .task_routing import verify_package
                verify_package(temp, plan.task_validation['projection'])
                log['task_route_verification'] = {'status': 'passed', 'scope': 'coalesced text/direct-format XML and exact non-target package parts'}
            if self.render_validation:
                from .advanced_writes import render_gate
                try:
                    targets = creation_plan or [(temp, output)]
                    reports = [render_gate(self.space.root, backup, staged, plan.task_validation['render_request'], approved_anchors=plan.task_validation['render_anchors']) if plan.task_validation is not None
                               else render_gate(self.space.root, backup, staged) for staged, _ in targets]
                    log['render_diff'] = reports[0] if len(reports) == 1 else reports
                except GuardError as error:
                    if hasattr(error, 'render_diff'):
                        log['render_diff'] = error.render_diff
                    raise
            if digest(src) != plan.source_sha256 or digest(backup) != plan.source_sha256:
                raise GuardError('source or backup changed before publication')
            # No overwrite, even on a race with an existing file. Hard-link exclusive publish.
            if creation_plan is None:
                os.link(temp, output)
                published_paths.append(output)
                published = True
                temp.unlink()
                log["status"] = "verified"
                log["output_sha256"] = digest(output)
            else:
                # A merge/split publishes every new document exclusively, then records them
                # together so a human rollback can remove exactly this job's outputs.
                published_paths = []
                for staged, destination in creation_plan:
                    os.link(staged, destination)
                    published_paths.append(destination)
                for staged, _ in creation_plan:
                    staged.unlink(missing_ok=True)
                published = True
                log["status"] = "verified"
                log["outputs"] = [str(path) for path in published_paths]
                log["outputs_sha256"] = [digest(path) for path in published_paths]
                log["output"] = str(published_paths[0])
                log["output_sha256"] = log["outputs_sha256"][0]
            log["diff"] = {"target": plan.target, "before": plan.before, "after": plan.after}
            (job / "audit.json").write_text(json.dumps(log, ensure_ascii=False, indent=2), encoding="utf-8")
            audited = True
            return log
        finally:
            temp.unlink(missing_ok=True)
            if not audited:
                log["status"] = "failed"
                # Publication can fail on the second split output: remove the first
                # even though the entire publish loop never reached `published = True`.
                for path in published_paths:
                    path.unlink(missing_ok=True)
                parts_directory = job / "parts"
                if parts_directory.is_dir():
                    for part in parts_directory.glob("*.docx"):
                        part.unlink(missing_ok=True)
                try:
                    (job / "audit.json").write_text(json.dumps(log, ensure_ascii=False, indent=2), encoding="utf-8")
                except OSError:
                    pass  # Even a full/unwritable disk cannot justify deleting the backup/source.

    def _creation_outputs(self, plan: Plan, src: Path, job: Path, temp: Path
                          ) -> tuple[list[tuple[Path, Path]], dict[str, Any]]:
        """Stage the new documents a merge/split creates, without touching any source.

        Returns ``([(staged file inside the audit job, final destination)], audit info)``.
        Every source is backed up first and re-hashed afterwards: a merge or split that
        changed one byte of an input would fail here instead of being published.
        """
        if isinstance(plan.request, DocxMerge):
            sources = _merge_sources(self.space, src, plan.request.sources)
            if src not in sources:
                raise GuardError("the authorized file must be one of the merge sources")
            before, after, _ = _merge_plan_value(sources)
            if before != plan.before:
                raise GuardError("merge inputs changed since preview")
            if after != plan.after:
                raise GuardError("merge preview no longer matches the inputs")
            backups = job / "sources"
            backups.mkdir()
            expected_digests = [item["sha256"] for item in json.loads(plan.before)["inputs"]]
            for index, (source, expected_digest) in enumerate(zip(sources, expected_digests), 1):
                source_backup = backups / f"{index:02d}-{source.name}"
                shutil.copy2(source, source_backup)
                if digest(source_backup) != expected_digest or digest(source) != expected_digest:
                    raise GuardError("merge source backup differs from confirmed preview")
            digests_before = expected_digests
            copy_report = _merge_docx_documents(sources, temp)
            if [digest(source) for source in sources] != digests_before:
                raise GuardError("a merge source changed while merging")
            reopened = _docx_signature(temp)
            if _signature_value(reopened["texts"], reopened["tables"],
                                reopened["sections"]) != json.loads(plan.after):
                raise GuardError("merged output does not match the preview")
            info = {"kind": "docx_merge", "source_backups": str(backups),
                    "fidelity_policy": dict(_B7_POLICY),
                    "output_naming": "single", "copy_report": copy_report,
                    "sources": [{"name": source.name, "sha256": source_digest}
                                for source, source_digest in zip(sources, digests_before)]}
            return [(temp, self._creation_destination(src, plan.id, None))], info

        before, after, titles, _ = _split_plan_value(src, plan.request.level)
        if before != plan.before:
            raise GuardError("split inputs changed since preview")
        if after != plan.after:
            raise GuardError("split preview no longer matches the inputs")
        source_digest = digest(src)
        document = Document(src)
        section = document.element.body.find(qn("w:sectPr"))
        parts = _split_parts(src, plan.request.level)
        staged: list[tuple[Path, Path]] = []
        texts: list[str] = []
        copy_reports: list[dict[str, Any]] = []
        directory = job / "parts"
        directory.mkdir()
        for index, (title, blocks) in enumerate(parts, 1):
            # Every destination is checked with the same guard before anything is written.
            destination = self.space.file(
                str(self._creation_destination(src, plan.id, index)), existing=False)
            part = directory / f"{index:02d}.docx"
            copy_reports.append(_write_split_part(blocks, section, part, document))
            reopened = Document(part)
            texts.extend(paragraph.text for paragraph in reopened.paragraphs if paragraph.text)
            staged.append((part, destination))
        if _signature_value(texts, 0, 0) != json.loads(plan.after):
            raise GuardError("split output does not match the preview")
        if digest(src) != source_digest:
            raise GuardError("the split source changed while splitting")
        info = {"kind": "docx_split", "split_source": src.name, "split_source_sha256": source_digest,
                "fidelity_policy": dict(_B7_POLICY), "part_copy_reports": copy_reports,
                "output_naming": "parts", "part_titles": titles, "part_count": len(staged),
                "copy_report": {"relationships_remapped": sum(report["relationships_remapped"]
                                                              for report in copy_reports),
                                "numbered_paragraphs": sum(report["numbered_paragraphs"]
                                                           for report in copy_reports)}}
        return staged, info

    @staticmethod
    def _creation_destination(src: Path, plan_id: str, part: int | None) -> Path:
        """The one naming rule merge/split outputs may use; rollback verifies it too."""
        if part is None:
            return src.with_name(f"{src.stem}.agent-{plan_id[:8]}{src.suffix}")
        return src.with_name(f"{src.stem}.part-{part:02d}.agent-{plan_id[:8]}{src.suffix}")

    def rollback(self, plan_id: str, *, confirmed: bool = False) -> dict:
        """Human-requested undo of a verified output copy; source and backup remain.

        Only our unchanged, audited output may be deleted. Never delete a changed or
        replacement file, even if it happens to occupy the published output path.
        """
        if confirmed is not True:
            raise GuardError("rollback requires a separate human confirmation")
        if not re.fullmatch(r"[0-9a-f]{32}", plan_id):
            raise GuardError("invalid plan identifier")
        job = self.audit / plan_id
        if job.is_symlink() or not job.is_dir() or not job.resolve(strict=True).is_relative_to(self.space.root):
            raise GuardError("audit job missing or redirected")
        record = job / "audit.json"
        if record.is_symlink():
            raise GuardError("audit record redirected")
        log = json.loads(record.read_text(encoding="utf-8"))
        if log.get("id") != plan_id or log.get("status") not in {"verified", "rolling_back"}:
            raise GuardError("only verified jobs can be rolled back")
        source = self.space.file(log["source"])
        recorded = log.get("outputs")
        if recorded:
            # A merge/split job publishes several new documents; every recorded path must
            # still match the one naming rule this module uses, or nothing is deleted.
            naming = log.get("output_naming")
            if naming == "parts":
                expected_outputs = [self._creation_destination(source, plan_id, index)
                                    for index in range(1, len(recorded) + 1)]
            elif naming == "single" and len(recorded) == 1:
                expected_outputs = [self._creation_destination(source, plan_id, None)]
            else:
                raise GuardError("audit output naming does not match original plan")
            if [str(path) for path in expected_outputs] != list(recorded):
                raise GuardError("audit output paths do not match original plan")
            digests = log.get("outputs_sha256") or []
            if len(digests) != len(expected_outputs):
                raise GuardError("audit output digests do not match original plan")
            if log.get("status") == "rolling_back" and not any(path.exists()
                                                               for path in expected_outputs):
                log["status"] = "rolled_back"
                record.write_text(json.dumps(log, ensure_ascii=False, indent=2), encoding="utf-8")
                return log
            outputs = []
            for path, recorded_digest in zip(expected_outputs, digests):
                current = self.space.file(str(path))
                if current != path or digest(current) != recorded_digest:
                    raise GuardError("output changed since publication; refusing deletion")
                outputs.append(current)
            log["status"] = "rolling_back"
            record.write_text(json.dumps(log, ensure_ascii=False, indent=2), encoding="utf-8")
            for path, recorded_digest in zip(outputs, digests):
                # Recheck immediately before removal, exactly as the single-output path does.
                if self.space.file(str(path)) != path or digest(path) != recorded_digest:
                    raise GuardError("output changed during rollback; refusing deletion")
                path.unlink()
            log["status"] = "rolled_back"
            record.write_text(json.dumps(log, ensure_ascii=False, indent=2), encoding="utf-8")
            return log
        suffix = source.suffix
        if log.get('kind') == 'document_export':
            details = json.loads(log['after'])
            suffix = {'pdf': '.pdf', 'csv': '.csv'}.get(details.get('format'))
            if suffix is None or source.suffix.lower() != {'.pdf': '.docx', '.csv': '.xlsx'}[suffix]:
                raise GuardError('invalid audited export format')
        expected = source.with_name(f"{source.stem}.agent-{plan_id[:8]}{suffix}")
        if log.get("output") != str(expected):
            raise GuardError("audit output path does not match original plan")
        if log.get("status") == "rolling_back" and not expected.exists():
            log["status"] = "rolled_back"
            record.write_text(json.dumps(log, ensure_ascii=False, indent=2), encoding="utf-8")
            return log
        output = self.space.file(log["output"])
        if output != expected or digest(output) != log.get("output_sha256"):
            raise GuardError("output changed since publication; refusing deletion")
        # Recheck immediately before removal; on shared hostile directories an OS
        # handle-based file lock would still be required to close replacement races.
        if self.space.file(str(output)) != expected or digest(output) != log["output_sha256"]:
            raise GuardError("output changed during rollback; refusing deletion")
        log["status"] = "rolling_back"
        record.write_text(json.dumps(log, ensure_ascii=False, indent=2), encoding="utf-8")
        output.unlink()
        log["status"] = "rolled_back"
        record.write_text(json.dumps(log, ensure_ascii=False, indent=2), encoding="utf-8")
        return log


# Maintenance is deliberately outside app.build_tools / the model dispatcher.
_AUDIT_BASE_FIELDS = {"id", "kind", "source", "source_sha256", "target", "before",
                      "after", "risks", "confirmation_required", "status", "backup", "output"}
_AUDIT_STATUSES = {"failed", "verified", "rolling_back", "rolled_back"}
_AUDIT_MAX_RECORD_BYTES = 4 * 1024 * 1024


def _audit_path(path: Path, *, directory: bool, missing: bool = False) -> None:
    """Check *every* lexical ancestor; resolve() alone hides junctions/reparse points."""
    if not path.is_absolute() or ".." in path.parts:
        raise GuardError("absolute non-traversing audit path required")
    for component in [*reversed(path.parents), path]:
        if component.name and (component.name != component.name.rstrip(" .")
                               or _RESERVED_DEVICE.fullmatch(component.name)):
            raise GuardError("unsafe audit path component")
        try:
            info = component.lstat()
        except FileNotFoundError:
            if component == path and missing:
                return
            raise GuardError("audit path missing") from None
        if (stat.S_ISLNK(info.st_mode)
                or getattr(info, "st_file_attributes", 0) & 0x400):
            raise GuardError("audit links/reparse points rejected")
        if component != path and not stat.S_ISDIR(info.st_mode):
            raise GuardError("ordinary audit ancestor directory required")
    if directory != stat.S_ISDIR(info.st_mode):
        raise GuardError("unexpected audit entry type")
    if not directory and not stat.S_ISREG(info.st_mode):
        raise GuardError("ordinary audit file required")


def _audit_root(authorized_root: Path) -> tuple[Workspace, Path]:
    raw = Path(os.path.abspath(authorized_root))
    _audit_path(raw, directory=True)
    space = Workspace(raw)
    audit = space.root / ".office_agent"
    _audit_path(audit, directory=True, missing=True)
    return space, audit


def _audit_record(space: Workspace, job: Path, *,
                  locked_payload: bytes | None = None) -> tuple[dict[str, Any], dict[str, Any]]:
    _audit_path(job, directory=True)
    if job.parent != space.root / ".office_agent" or not re.fullmatch(r"[0-9a-f]{32}", job.name):
        raise GuardError("not a current audit job directory")
    record = job / "audit.json"
    _audit_path(record, directory=False)
    if space.file(str(record)) != record:
        raise GuardError("audit record escaped workspace")
    first = record.stat()
    if first.st_size > _AUDIT_MAX_RECORD_BYTES or first.st_nlink != 1:
        raise GuardError("oversized or hard-linked audit record retained")
    payload = record.read_bytes() if locked_payload is None else locked_payload
    last = record.stat()
    identity = lambda s: (s.st_dev, s.st_ino, s.st_size, s.st_mtime_ns, s.st_ctime_ns)
    if identity(first) != identity(last):
        raise GuardError("audit record changed while reading")
    def unique_pairs(pairs):
        result = {}
        for key, value in pairs:
            if key in result:
                raise GuardError("duplicate audit JSON key")
            result[key] = value
        return result
    log = json.loads(payload, object_pairs_hook=unique_pairs)
    if not isinstance(log, dict) or not _AUDIT_BASE_FIELDS <= log.keys():
        raise GuardError("unknown/legacy audit format")
    if log["id"] != job.name or log["status"] not in _AUDIT_STATUSES:
        raise GuardError("audit identity/status is not current format")
    if any(not isinstance(log[key], str) for key in
           ("kind", "source", "source_sha256", "target", "before", "after", "backup", "output")):
        raise GuardError("invalid audit string fields")
    if (not log["kind"] or not re.fullmatch(r"[0-9a-f]{64}", log["source_sha256"])
            or type(log["confirmation_required"]) is not bool
            or not isinstance(log["risks"], list)
            or any(not isinstance(item, str) for item in log["risks"])):
        raise GuardError("invalid audit plan fields")
    evidence = {"sha256": hashlib.sha256(payload).hexdigest(), "mtime_ns": last.st_mtime_ns,
                "size_bytes": last.st_size, "device": last.st_dev, "inode": last.st_ino,
                "ctime_ns": last.st_ctime_ns}
    return log, evidence


def _audit_no_write(space: Workspace, job: Path, log: dict[str, Any],
                    evidence: dict[str, Any], cutoff_ns: int) -> None:
    # Current writer does not emit a cancelled record; unfamiliar statuses are retained.
    if log["status"] != "failed":
        raise GuardError("verified/rollback/non-failed audit always retained")
    if set(log) != _AUDIT_BASE_FIELDS:
        raise GuardError("additional write/publication/unknown evidence retained")
    if evidence["mtime_ns"] >= cutoff_ns:
        raise GuardError("audit record is not older than cutoff")
    if sorted(item.name for item in job.iterdir()) != ["audit.json"]:
        raise GuardError("backup/working/source/other remnants retained")
    source_path = Path(log["source"])
    _audit_path(source_path, directory=False)
    source = space.file(log["source"])
    if source != source_path or source.parent == job or source.is_relative_to(job.parent):
        raise GuardError("source path is not an ordinary external-to-audit document")
    if source.suffix.lower() not in {".docx", ".xlsx", ".doc", ".xls"}:
        raise GuardError("unknown source format retained")
    expected_backup = job / ("backup" + source.suffix)
    expected_output = source.with_name(f"{source.stem}.agent-{job.name[:8]}{source.suffix}")
    if log["backup"] != str(expected_backup) or log["output"] != str(expected_output):
        raise GuardError("backup/output naming not provable")
    # No source, output, backup or rollback data is ever removed by maintenance.
    _audit_path(expected_backup, directory=False, missing=True)
    _audit_path(expected_output, directory=False, missing=True)
    if expected_backup.exists() or expected_output.exists():
        raise GuardError("backup or output exists; audit retained")
    if digest(source) != log["source_sha256"]:
        raise GuardError("source changed; no-write failure cannot be proven")


def audit_summary(authorized_root: Path) -> dict[str, Any]:
    """Read-only inventory. Times are audit.json mtime (NOT operation timestamps).

    Occupancy counts regular files, including backups, without following any links.
    Inaccessible or unsafe entries are listed; size is explicitly partial in that case.
    """
    space, audit = _audit_root(authorized_root)
    result: dict[str, Any] = {"audit_directory": str(audit), "status_counts": {},
        "oldest": None, "newest": None, "time_basis": "audit.json filesystem mtime UTC",
        "size_bytes": 0, "size_complete": True, "inventory_complete": True,
        "scanned_entries": 0, "scan_limits": {"entries": 20000, "depth": 32},
        "failures": [], "damaged": [], "records": []}
    if not audit.exists():
        return result
    # Inventory must not recurse without a bound or crash if the root disappears.
    # Reuse one root snapshot for records instead of a second unguarded iterdir().
    max_entries, max_depth = result["scan_limits"]["entries"], result["scan_limits"]["depth"]
    scanned = 0

    def incomplete(path, reason):
        result["size_complete"] = False
        result["damaged"].append({"path": str(path), "reason": reason})

    def entries(directory):
        nonlocal scanned
        _audit_path(directory, directory=True)
        found = []
        with os.scandir(directory) as iterator:
            for entry in iterator:
                if scanned >= max_entries:
                    incomplete(directory, "inventory entry limit reached; totals are partial")
                    if directory == audit:
                        result["inventory_complete"] = False
                    break
                scanned += 1
                found.append(Path(entry.path))
        return sorted(found, key=lambda p: p.name)

    try:
        root_entries = entries(audit)
    except (OSError, GuardError) as exc:
        incomplete(audit, str(exc))
        result["inventory_complete"] = False
        return result
    stack = [(root_entries, 0)]
    while stack:
        children, depth = stack.pop()
        for item in children:
            try:
                info = item.lstat()
                is_directory = stat.S_ISDIR(info.st_mode)
                _audit_path(item, directory=is_directory)
                if is_directory:
                    if depth >= max_depth:
                        incomplete(item, "inventory depth limit reached; totals are partial")
                    else:
                        stack.append((entries(item), depth + 1))
                else:
                    result["size_bytes"] += info.st_size
            except (OSError, GuardError) as exc:
                incomplete(item, str(exc))
    result["scanned_entries"] = scanned
    times = []
    for job in root_entries:
        # Maintenance journals are files, not jobs; never parse them as candidates.
        if job.name.startswith("retention-") and job.suffix == ".jsonl":
            continue
        try:
            log, evidence = _audit_record(space, job)
            item = {"id": job.name, "status": log["status"], **evidence}
            result["records"].append(item)
            counts = result["status_counts"]
            counts[log["status"]] = counts.get(log["status"], 0) + 1
            times.append(evidence["mtime_ns"])
            if log["status"] == "failed":
                result["failures"].append(item)
        except (OSError, ValueError, TypeError) as exc:
            result["damaged"].append({"path": str(job), "reason": str(exc)})
    if times:
        utc = lambda ns: datetime.fromtimestamp(ns / 1e9, timezone.utc).isoformat()
        result["oldest"], result["newest"] = utc(min(times)), utc(max(times))
    return result


def _audit_locked_delete(space: Workspace, job: Path, expected: dict[str, Any],
                         cutoff_ns: int, journal) -> None:
    """Windows handle deletion: no pathname unlink race. Elsewhere fail closed.

    Directory handles disallow rename/delete of every ancestor; the record handle
    disallows competing writers/deleters until final revalidation and disposition.
    Only audit.json is disposed; no recursive delete and no directory removal.
    """
    if os.name != "nt":
        raise GuardError("safe handle-based cleanup currently requires Windows; retained")
    import ctypes
    from ctypes import wintypes
    kernel = ctypes.WinDLL("kernel32", use_last_error=True)
    create = kernel.CreateFileW
    create.argtypes = [wintypes.LPCWSTR, wintypes.DWORD, wintypes.DWORD,
                       wintypes.LPVOID, wintypes.DWORD, wintypes.DWORD, wintypes.HANDLE]
    create.restype = wintypes.HANDLE
    close = kernel.CloseHandle
    close.argtypes = [wintypes.HANDLE]
    close.restype = wintypes.BOOL
    disposition = kernel.SetFileInformationByHandle
    disposition.argtypes = [wintypes.HANDLE, ctypes.c_int, wintypes.LPVOID, wintypes.DWORD]
    disposition.restype = wintypes.BOOL
    handles = []
    def pin(path, access, flags):
        _audit_path(path, directory=bool(flags & 0x02000000))
        handle = create(str(path), access, 1, None, 3, flags | 0x00200000, None)
        if handle == ctypes.c_void_p(-1).value:
            raise GuardError(f"cannot safely lock audit path: Windows error {ctypes.get_last_error()}")
        handles.append(handle)
        _audit_path(path, directory=bool(flags & 0x02000000))
        return handle
    try:
        for ancestor in [*reversed(job.parents), job]:
            pin(ancestor, 0x80, 0x02000000)  # FILE_READ_ATTRIBUTES, BACKUP_SEMANTICS
        record_handle = pin(job / "audit.json", 0x80000000 | 0x10000, 0)
        read_file = kernel.ReadFile
        read_file.argtypes = [wintypes.HANDLE, wintypes.LPVOID, wintypes.DWORD,
                             ctypes.POINTER(wintypes.DWORD), wintypes.LPVOID]
        read_file.restype = wintypes.BOOL
        buffer = ctypes.create_string_buffer(_AUDIT_MAX_RECORD_BYTES + 1)
        count = wintypes.DWORD()
        if not read_file(record_handle, buffer, len(buffer), ctypes.byref(count), None):
            raise GuardError("locked audit record could not be read")
        payload = buffer.raw[:count.value]
        log, evidence = _audit_record(space, job, locked_payload=payload)
        if evidence != expected:
            raise GuardError("audit identity/content/time changed after preview; retained")
        _audit_no_write(space, job, log, evidence, cutoff_ns)
        journal({"event": "delete_intent", "id": job.name, "evidence": evidence,
                 "record": log})  # Preserve provenance even if a new artifact appears concurrently.
        # Revalidate after durable intent; locked record cannot change status/mtime/content.
        current, current_evidence = _audit_record(space, job, locked_payload=payload)
        if current_evidence != expected:
            raise GuardError("audit changed before disposal; retained")
        _audit_no_write(space, job, current, current_evidence, cutoff_ns)
        flag = ctypes.c_ubyte(1)  # FILE_DISPOSITION_INFO.DeleteFile is BOOLEAN.
        if not disposition(record_handle, 4, ctypes.byref(flag), ctypes.sizeof(flag)):
            raise GuardError(f"safe audit disposal refused: Windows error {ctypes.get_last_error()}")
    finally:
        for handle in reversed(handles):
            close(handle)


def retain_audits(authorized_root: Path, *, older_than_days: int = 90,
                  apply: bool = False) -> dict[str, Any]:
    """Explicit retention preview/apply, 1..3650 days, never exposed to the model.

    Deletes only proven no-write failed audit.json records, retaining empty job
    directories and all artifacts. Unknown/cancelled/verified/rollback records stay.
    Failed jobs with their normal backups are intentionally NOT reclaimable.
    """
    if type(older_than_days) is not int or not 1 <= older_than_days <= 3650:
        raise GuardError("older_than_days must be an integer between 1 and 3650")
    if type(apply) is not bool:
        raise GuardError("apply must be an explicit boolean")
    space, audit = _audit_root(authorized_root)
    cutoff_ns = time.time_ns() - older_than_days * 86400 * 1_000_000_000
    summary = audit_summary(space.root)
    result: dict[str, Any] = {"dry_run": not apply, "older_than_days": older_than_days,
        "cutoff_ns": cutoff_ns, "time_basis": summary["time_basis"], "candidates": [],
        "retained": list(summary["damaged"]), "deleted": [], "journal": None}
    for item in summary["records"]:
        job = audit / item["id"]
        try:
            log, evidence = _audit_record(space, job)
            _audit_no_write(space, job, log, evidence, cutoff_ns)
            result["candidates"].append({"id": job.name, **evidence})
        except (OSError, ValueError, TypeError) as exc:
            result["retained"].append({"id": job.name, "reason": str(exc)})
    if not apply or not result["candidates"]:
        return result
    if os.name != "nt":
        for item in result["candidates"]:
            result["retained"].append({"id": item["id"], "reason": "Windows handle locking required"})
        return result
    # Pin audit ancestry before creating the exclusive maintenance journal too.
    # A journal is never followed or reused: unique file, exclusive create, fsync.
    _audit_path(audit, directory=True)
    journal_path = audit / f"retention-{uuid.uuid4().hex}.jsonl"
    # Opening the journal under a raced directory could otherwise create outside root.
    import ctypes
    from ctypes import wintypes
    kernel = ctypes.WinDLL("kernel32", use_last_error=True)
    create = kernel.CreateFileW
    create.argtypes = [wintypes.LPCWSTR, wintypes.DWORD, wintypes.DWORD, wintypes.LPVOID,
                       wintypes.DWORD, wintypes.DWORD, wintypes.HANDLE]
    create.restype = wintypes.HANDLE
    close = kernel.CloseHandle
    close.argtypes = [wintypes.HANDLE]
    close.restype = wintypes.BOOL
    locks = []
    try:
        for directory in [*reversed(audit.parents), audit]:
            _audit_path(directory, directory=True)
            handle = create(str(directory), 0x80, 1, None, 3, 0x02200000, None)
            if handle == ctypes.c_void_p(-1).value:
                raise GuardError("audit directory cannot be safely locked for journal")
            locks.append(handle)
            _audit_path(directory, directory=True)
        with journal_path.open("x", encoding="utf-8") as stream:
            def journal(event):
                stream.write(json.dumps({"utc": datetime.now(timezone.utc).isoformat(), **event},
                                        ensure_ascii=False) + "\n")
                stream.flush()
                os.fsync(stream.fileno())
            result["journal"] = str(journal_path)
            journal({"event": "retention_start", "older_than_days": older_than_days,
                     "cutoff_ns": cutoff_ns, "policy": "failed-audit-only-no-remnants-v1"})
            for item in result["candidates"]:
                job = audit / item["id"]
                try:
                    _audit_locked_delete(space, job,
                        {key: value for key, value in item.items() if key != "id"}, cutoff_ns, journal)
                except (OSError, ValueError, TypeError) as exc:
                    kept = {"id": job.name, "reason": str(exc)}
                    result["retained"].append(kept)
                    journal({"event": "retained", **kept})
                else:
                    result["deleted"].append(job.name)
                    journal({"event": "deleted", "id": job.name})
            journal({"event": "retention_complete", "deleted": result["deleted"]})
    finally:
        for handle in reversed(locks):
            close(handle)
    return result
