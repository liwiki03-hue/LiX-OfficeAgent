"""Focused B7 OOXML structure checks; no Office, COM, or rendering."""
import json
from pathlib import Path
import struct
import zlib

import pytest
from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from office_agent.writes import DocxMerge, DocxSplit, Editor, GuardError


def _ref(parent, tag, value):
    node = OxmlElement(tag)
    node.set(qn("w:val"), value)
    parent.append(node)
    return node


def _source(path, bold=False, chapters=False):
    doc = Document()
    char = doc.styles.add_style("B7 Character", WD_STYLE_TYPE.CHARACTER)
    style = doc.styles.add_style("B7 Fancy", WD_STYLE_TYPE.PARAGRAPH)
    style.base_style = doc.styles["Normal"]
    style.font.bold = bold
    style.next_paragraph_style = style
    _ref(style.element, "w:link", char.style_id)
    _ref(char.element, "w:link", style.style_id)
    numpr = OxmlElement("w:numPr")
    _ref(numpr, "w:numId", "1")
    style.element.get_or_add_pPr().append(numpr)
    # Exercise numbering -> style references as well as style -> numbering.
    abstract = doc.part.numbering_part.element.find(qn("w:abstractNum"))
    _ref(abstract, "w:styleLink", style.style_id)
    for i in range(2 if chapters else 1):
        if chapters:
            doc.add_heading(f"Chapter {i + 1}", level=1)
        doc.add_paragraph(f"{path.stem} item {i}", style=style)
    doc.save(path)


def _assert_list_graph(doc, paragraph):
    style = paragraph.style
    assert style.name.startswith("B7 Fancy [B7 source ")
    assert style.base_style.style_id.startswith("B7s")
    assert style.next_paragraph_style.style_id == style.style_id
    link = style.element.find(qn("w:link")).get(qn("w:val"))
    linked_style = next(s for s in doc.styles if s.style_id == link)
    assert linked_style.element.find(qn("w:link")).get(qn("w:val")) == style.style_id
    numid = style.element.find("./" + qn("w:pPr") + "/" + qn("w:numPr") + "/" + qn("w:numId")).get(qn("w:val"))
    numbering = doc.part.numbering_part.element
    num = next(n for n in numbering.findall(qn("w:num")) if n.get(qn("w:numId")) == numid)
    abstractid = num.find(qn("w:abstractNumId")).get(qn("w:val"))
    assert any(n.get(qn("w:abstractNumId")) == abstractid for n in numbering.findall(qn("w:abstractNum")))
    assert any(n.get(qn("w:val")) == style.style_id for n in numbering.iter(qn("w:styleLink")))
    return numid


def test_merge_isolates_same_name_styles_and_inherited_numbering(tmp_path):
    first, second = tmp_path / "a.docx", tmp_path / "b.docx"
    _source(first)
    _source(second, bold=True)
    original = [p.read_bytes() for p in (first, second)]
    editor = Editor(tmp_path)
    plan = editor.plan(DocxMerge(path=str(first), sources=[first.name, second.name]))
    assert json.loads(plan.before)["fidelity_policy"]["header_footer_footnotes_endnotes_comments"].startswith("refuse")
    assert json.loads(plan.before)["inputs"][0]["numbered"] == 1
    with pytest.raises(GuardError, match="human confirmation"):
        editor.execute(plan.id)
    assert not editor.audit.exists()
    log = editor.execute(plan.id, confirmed=True)
    output = Path(log["output"])
    doc = Document(output)
    paragraphs = [p for p in doc.paragraphs if p.text]
    assert paragraphs[0].style.style_id != paragraphs[1].style.style_id
    assert paragraphs[0].style.font.bold is False
    assert paragraphs[1].style.font.bold is True
    assert _assert_list_graph(doc, paragraphs[0]) != _assert_list_graph(doc, paragraphs[1])
    assert log["copy_report"]["numbered_paragraphs"] == 2
    assert "visual fidelity not verified" in log["fidelity_policy"]["verification"]
    assert [p.read_bytes() for p in (first, second)] == original
    editor.rollback(plan.id, confirmed=True)
    assert not output.exists()


def test_split_keeps_custom_style_numbering_and_source_bytes(tmp_path):
    source = tmp_path / "chapters.docx"
    _source(source, chapters=True)
    before = source.read_bytes()
    editor = Editor(tmp_path)
    plan = editor.plan(DocxSplit(path=str(source)))
    log = editor.execute(plan.id, confirmed=True)
    assert log["copy_report"]["numbered_paragraphs"] == 2
    for output in log["outputs"]:
        doc = Document(output)
        _assert_list_graph(doc, doc.paragraphs[1])
    assert source.read_bytes() == before
    editor.rollback(plan.id, confirmed=True)
    assert all(not Path(p).exists() for p in log["outputs"])


@pytest.mark.parametrize("bad", ["header", "style", "numbering", "defaults"])
def test_merge_split_refuse_unmigratable_inputs_before_outputs(tmp_path, bad):
    source, other = tmp_path / "bad.docx", tmp_path / "other.docx"
    _source(source, chapters=True)
    _source(other)
    doc = Document(source)
    if bad == "header":
        doc.sections[0].header.paragraphs[0].text = "must not disappear"
    elif bad == "style":
        _ref(doc.styles["B7 Fancy"].element, "w:basedOn", "MissingStyle")
        # Replace the pre-existing basedOn too: every reference is checked.
    elif bad == "numbering":
        doc.styles["B7 Fancy"].element.find("./" + qn("w:pPr") + "/" + qn("w:numPr") + "/" + qn("w:numId")).set(qn("w:val"), "99999")
    else:
        defaults = doc.styles.element.find(qn("w:docDefaults"))
        defaults.set(qn("w:testMismatch"), "1")
    doc.save(source)
    before = source.read_bytes()
    editor = Editor(tmp_path)
    for request in (DocxMerge(path=str(other), sources=[other.name, source.name]), DocxSplit(path=str(source))):
        with pytest.raises(GuardError, match="B7"):
            editor.plan(request)
    assert source.read_bytes() == before
    assert not editor.audit.exists()
    assert not list(tmp_path.glob("*.agent-*"))


def test_failed_second_split_publication_removes_first_and_staging(tmp_path, monkeypatch):
    import office_agent.writes as writes
    source = tmp_path / "chapters.docx"
    _source(source, chapters=True)
    editor = Editor(tmp_path)
    plan = editor.plan(DocxSplit(path=str(source)))
    real_link = writes.os.link
    calls = []
    def failing_link(staged, destination):
        calls.append(destination)
        if len(calls) == 2:
            raise OSError("injected second publication failure")
        return real_link(staged, destination)
    monkeypatch.setattr(writes.os, "link", failing_link)
    with pytest.raises(OSError, match="second publication"):
        editor.execute(plan.id, confirmed=True)
    assert not list(tmp_path.glob("*.agent-*"))
    job = editor.audit / plan.id
    assert not list((job / "parts").glob("*.docx"))
    assert (job / "backup.docx").read_bytes() == source.read_bytes()
    assert json.loads((job / "audit.json").read_text(encoding="utf-8"))["status"] == "failed"


def _png(rgb):
    def chunk(kind, data):
        return struct.pack(">I", len(data)) + kind + data + struct.pack(">I", zlib.crc32(kind + data))
    return (b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", struct.pack(">IIBBBBB", 1, 1, 8, 2, 0, 0, 0))
            + chunk(b"IDAT", zlib.compress(b"\x00" + bytes(rgb))) + chunk(b"IEND", b""))


def test_merge_does_not_reuse_relationship_ids_or_image_partnames(tmp_path):
    sources, expected = [], []
    for index, color in enumerate(((255, 0, 0), (0, 0, 255))):
        image = tmp_path / f"{index}.png"
        image.write_bytes(_png(color))
        doc = Document()
        doc.add_picture(str(image))
        path = tmp_path / f"{index}.docx"
        doc.save(path)
        sources.append(path)
        expected.append(image.read_bytes())
    editor = Editor(tmp_path)
    plan = editor.plan(DocxMerge(path=str(sources[0]), sources=[p.name for p in sources]))
    log = editor.execute(plan.id, confirmed=True)
    doc = Document(log["output"])
    actual = []
    for node in doc.element.body.iter(qn("a:blip")):
        actual.append(doc.part.rels[node.get(qn("r:embed"))].target_part.blob)
    assert actual == expected
    assert log["copy_report"]["relationships_remapped"] == 2


def test_merge_checks_non_anchor_source_risks_before_output(tmp_path, monkeypatch):
    import office_agent.documents as documents
    anchor, risky = tmp_path / "safe.docx", tmp_path / "risky.docx"
    _source(anchor)
    _source(risky)
    original = documents.inspect_file
    def inspect(path):
        report = original(path)
        if Path(path) == risky:
            report["risks"] = ["unsupported_second_source_feature"]
        return report
    monkeypatch.setattr(documents, "inspect_file", inspect)
    editor = Editor(tmp_path)
    with pytest.raises(GuardError, match="B7 merge source.*risky.docx"):
        editor.plan(DocxMerge(path=str(anchor), sources=[anchor.name, risky.name]))
    assert not editor.audit.exists()
    assert not list(tmp_path.glob("*.agent-*"))
