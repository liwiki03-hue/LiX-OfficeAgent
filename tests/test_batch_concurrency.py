"""Bounded sibling reads: ordered, guarded, fail-isolated, and never writing files."""
import threading
from pathlib import Path

from docx import Document

from office_agent import documents
from office_agent.writes import Workspace


def test_batch_read_workers_overlap_but_results_keep_input_order(tmp_path, monkeypatch):
    for name in ("a.docx", "b.docx", "c.docx"):
        doc = Document()
        doc.add_paragraph("Synthetic " + name)
        doc.save(tmp_path / name)
    (tmp_path / "broken.docx").write_bytes(b"not an OOXML package")
    before = {path.name: path.read_bytes() for path in tmp_path.iterdir()}
    admitted = []
    workspace = Workspace(tmp_path)

    def validate(raw):
        admitted.append(Path(raw).name)
        return workspace.file(raw)

    rendezvous = threading.Barrier(3)
    c_done, b_done = threading.Event(), threading.Event()
    finish_order = []
    thread_ids = set()
    lock = threading.Lock()

    def controlled_read(candidate, *, operation, query, per_file):
        assert operation == "summary" and query is None and per_file == 20
        with lock:
            thread_ids.add(threading.get_ident())
        rendezvous.wait(timeout=5)  # Serial implementation fails; never hang the test.
        if candidate.name == "a.docx":
            assert b_done.wait(5)
        elif candidate.name == "b.docx":
            assert c_done.wait(5)
        with lock:
            finish_order.append(candidate.name)
        if candidate.name == "c.docx":
            c_done.set()
        elif candidate.name == "b.docx":
            b_done.set()
            raise ValueError("synthetic read failure")
        return {"file": candidate.name, "kind": "word", "preview": []}

    monkeypatch.setattr(documents, "_batch_file_report", controlled_read)
    report = documents.batch_document_report(tmp_path / "a.docx", validate=validate)
    assert admitted == ["a.docx", "b.docx", "broken.docx", "c.docx"]
    assert len(thread_ids) == 3 and finish_order == ["c.docx", "b.docx", "a.docx"]
    assert [entry["file"] for entry in report["files"]] == ["a.docx", "c.docx"]
    assert [entry["file"] for entry in report["errors"]] == ["b.docx"]
    assert [entry["file"] for entry in report["skipped"]] == ["broken.docx"]
    assert (report["file_count"], report["error_count"], report["skipped_count"]) == (2, 1, 1)
    assert report["wrote_nothing"] is True and report["truncated"] is False
    assert {path.name: path.read_bytes() for path in tmp_path.iterdir()} == before
