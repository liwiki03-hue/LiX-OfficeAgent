from pathlib import Path
import zipfile

from docx import Document
from docx.shared import Pt
import pytest

from office_agent import fidelity
from office_agent.diagram_process import OwnedRun
from office_agent.fidelity import (FidelityError, compare_documents, compare_pdf_snapshots,
                                   compare_structure, package_projection, pdf_snapshot)


def pdf_bytes(*, text="Hello", font_size=12, pages=1, width=200):
    stream = f"BT /F1 {font_size} Tf 20 100 Td ({text}) Tj ET".encode()
    objects = [b"<< /Type /Catalog /Pages 2 0 R >>",
               f"<< /Type /Pages /Kids [{' '.join(f'{5+i} 0 R' for i in range(pages))}] /Count {pages} >>".encode(),
               b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
               f"<< /Length {len(stream)} >>\nstream\n".encode() + stream + b"\nendstream"]
    objects.extend(f"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 {width} 200] /Resources << /Font << /F1 3 0 R >> >> /Contents 4 0 R >>".encode()
                   for _ in range(pages))
    data = b"%PDF-1.4\n"
    offsets = []
    for index, obj in enumerate(objects, start=1):
        offsets.append(len(data))
        data += f"{index} 0 obj\n".encode() + obj + b"\nendobj\n"
    start = len(data)
    data += f"xref\n0 {len(objects)+1}\n0000000000 65535 f \n".encode()
    data += b"".join(f"{offset:010} 00000 n \n".encode() for offset in offsets)
    return data + f"trailer\n<< /Size {len(objects)+1} /Root 1 0 R >>\nstartxref\n{start}\n%%EOF\n".encode()


@pytest.fixture
def pair(tmp_path):
    first, second = tmp_path / "first.docx", tmp_path / "second.docx"
    doc = Document()
    doc.add_paragraph("Hello")
    doc.save(first)
    second.write_bytes(first.read_bytes())
    return first, second


def test_structural_projection_and_style_only_change(pair):
    first, second = pair
    assert compare_structure(first, second)["equivalent"] is True
    doc = Document(second)
    doc.paragraphs[0].runs[0].font.size = Pt(30)
    doc.save(second)
    result = compare_structure(first, second)
    assert result["equivalent"] is False
    assert "word/document.xml" in result["changed_parts"]
    assert result["truncated"] is False


@pytest.mark.parametrize("options,text_equal,pixels_equal,pages_equal", [
    ({}, True, True, True),
    ({"text": "Other"}, False, False, True),
    ({"font_size": 25}, True, False, True),
    ({"pages": 2}, False, False, False),
])
def test_real_pdf_text_and_pixels_are_independent(tmp_path, options, text_equal, pixels_equal, pages_equal):
    left, right = tmp_path / "left.pdf", tmp_path / "right.pdf"
    left.write_bytes(pdf_bytes())
    right.write_bytes(pdf_bytes(**options))
    original = left.read_bytes(), right.read_bytes()
    result = compare_pdf_snapshots(pdf_snapshot(left), pdf_snapshot(right))
    assert result["text_equal"] is text_equal
    assert result["pixels_equal"] is pixels_equal
    assert result["page_count_equal"] is pages_equal
    assert result["equivalent"] is (text_equal and pixels_equal and pages_equal)
    assert (left.read_bytes(), right.read_bytes()) == original
    assert result["dpi"] == 144


@pytest.mark.parametrize("data,code", [(b"invalid", "invalid_pdf"),
                                       (pdf_bytes(pages=41), "page_count_limit"),
                                       (pdf_bytes(width=50000), "page_pixel_limit")])
def test_pdf_limits_fail_closed(tmp_path, data, code):
    path = tmp_path / "input.pdf"
    path.write_bytes(data)
    with pytest.raises(FidelityError, match=code):
        pdf_snapshot(path)
    assert path.read_bytes() == data


def test_runtime_missing_is_not_render_equivalence(pair, tmp_path):
    first, second = pair
    result = compare_documents(tmp_path, first, second, executable=tmp_path / "missing.exe")
    assert result["structure"]["equivalent"] is True
    assert result["render"] == {"status": "unavailable", "equivalent": None, "code": "libreoffice_unavailable"}
    assert result["inputs_unchanged"] is True
    assert result["processes"] == []


def test_timeout_never_returns_success(pair, tmp_path, monkeypatch):
    executable = tmp_path / "runtime.exe"
    executable.write_bytes(b"fake; never executed")
    monkeypatch.setattr(fidelity, "_run_owned", lambda *a, **k: OwnedRun("render_timeout", None, 12, .1, True))
    result = compare_documents(tmp_path, *pair, executable=executable, timeout_ms=100)
    assert result["render"]["code"] == "worker_timeout"
    assert result["render"]["equivalent"] is None
    assert result["inputs_unchanged"] is True
    assert result["processes"][0]["empty_confirmed"] is True


def test_success_exit_without_pdf_is_failure(pair, tmp_path, monkeypatch):
    executable = tmp_path / "runtime.exe"
    executable.write_bytes(b"fake; never executed")
    commands = []

    def run(command, **kwargs):
        commands.append(command)
        assert command[2] == "--convert"
        assert (Path(command[4]) / "profile" / "user" / "registrymodifications.xcu").is_file()
        copied = Path(command[3])
        assert copied.read_bytes() == pair[0].read_bytes()
        assert copied not in pair
        return OwnedRun("process_completed", 0, 12, .1, True)

    monkeypatch.setattr(fidelity, "_run_owned", run)
    result = compare_documents(tmp_path, *pair, executable=executable)
    assert result["render"]["status"] == "error"
    assert result["render"]["equivalent"] is None
    assert len(commands) == 1
    assert not Path(commands[0][3]).exists()


def test_external_relationship_refused_before_launch(pair, tmp_path, monkeypatch):
    first, second = pair
    with zipfile.ZipFile(first, "a") as archive:
        archive.writestr("custom/_rels/unsafe.rels", '<Relationships><Relationship TargetMode="External" Target="https://example.invalid/"/></Relationships>')
    with pytest.raises(FidelityError, match="external_relationship_rejected"):
        package_projection(first, for_render=True)
    executable = tmp_path / "runtime.exe"
    executable.write_bytes(b"fake; never executed")
    monkeypatch.setattr(fidelity, "_run_owned", lambda *a, **k: pytest.fail("must not launch"))
    report = compare_documents(tmp_path, first, second, executable=executable)
    assert report["render"]["code"] == "external_relationship_rejected"
    assert report["inputs_unchanged"] is True


def test_package_entity_and_oversize_refused(pair, monkeypatch):
    first, _ = pair
    with zipfile.ZipFile(first, "a") as archive:
        archive.writestr("custom/unsafe.xml", '<!DOCTYPE doc [<!ENTITY x "expanded">]><doc>&x;</doc>')
    with pytest.raises(FidelityError, match="xml_entities_rejected"):
        package_projection(first)
    monkeypatch.setattr(fidelity, "MAX_FILE_BYTES", 10)
    with pytest.raises(FidelityError, match="file_size_limit"):
        package_projection(first)


def test_unauthorized_input_is_refused(pair, tmp_path):
    root = tmp_path / "authorized"
    root.mkdir()
    with pytest.raises(ValueError, match="outside authorized"):
        compare_documents(root, *pair)
