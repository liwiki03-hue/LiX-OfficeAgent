import base64
import json
import tempfile
import zipfile
import unittest
from pathlib import Path
from unittest.mock import patch

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from pydantic import ValidationError
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font
from openpyxl.chart import BarChart, LineChart, PieChart, Reference
from openpyxl.workbook.defined_name import DefinedName

from office_agent.writes import (DocxMerge, DocxSplit, Editor, ExcelBarChart, ExcelChart, ExcelFormulaFill,
                                 ExcelRangeSet, ExcelRangeTrim, ExcelSet,
                                 ExcelTrim, GuardError, LegacyCellSet, LegacyWordReplace, WordBatchReplace,
                                 WordMergedCellText, WordRevisionDecision,
                                 _merge_projection,
                                 WordFormat, WordReplace, WordTableCellReplace, WordTableRow, Workspace, digest)
from office_agent.documents import read_word_table_structure


_OLE_HEADER = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"
_PIXEL_PNG = base64.b64decode(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/lZkAAAAASUVORK5CYII=")


class FakeLegacyBackend:
    """In-memory .doc/.xls double: no Office involved, deterministic and inspectable.

    The fake reads and writes the audit job's copies exactly like the real COM backend does,
    so the chain under test is the production one; only the Office instance is replaced.
    """

    def __init__(self, *, available=True, reason="COM 自动化被权限或策略拒绝", fail_on=None,
                 corrupt_paragraph=None, corrupt_cell=None):
        self._available = available
        self._reason = reason
        self.fail_on = fail_on
        self.corrupt_paragraph = corrupt_paragraph
        self.corrupt_cell = corrupt_cell
        self.probes: list[str] = []
        self.reads: list[str] = []
        self.writes: list[str] = []

    def probe(self, application):
        self.probes.append(application)
        return {"application": application, "available": self._available, "stage": "startup",
                "probe": "blank_document_save_reopen",
                "reason": "" if self._available else self._reason}

    @staticmethod
    def _payload(path):
        data = Path(path).read_bytes()
        if data[:8] != _OLE_HEADER:
            raise AssertionError("the fake backend only reads OLE fixtures")
        return data[8:]

    def read_word_paragraphs(self, path, *, limit):
        self.reads.append(str(path))
        return [line.decode("utf-8") for line in self._payload(path).split(b"\n")][:limit]

    def replace_word_paragraph(self, path, *, index, old, new, expected_count):
        if self.fail_on == "write":
            raise RuntimeError("simulated COM failure during write")
        self.writes.append(str(path))
        lines = self._payload(path).split(b"\n")
        body = lines[index].decode("utf-8")
        if body.count(old) != expected_count:
            raise ValueError("count mismatch")
        lines[index] = body.replace(old, new).encode("utf-8")
        if self.corrupt_paragraph is not None and self.corrupt_paragraph != index:
            lines[self.corrupt_paragraph] = b"CORRUPTED BY THE BACKEND"
        Path(path).write_bytes(_OLE_HEADER + b"\n".join(lines))
        return self.read_word_paragraphs(path, limit=index + 1)[index]

    def _cells(self, path):
        return json.loads(self._payload(path).decode("utf-8"))

    def read_excel_cell(self, path, *, sheet, cell):
        self.reads.append(str(path))
        return self._cells(path).get(sheet, {}).get(cell)

    def read_excel_region(self, path, *, sheet, first_cell, last_cell):
        return dict(self._cells(path).get(sheet, {}))

    def set_excel_cell(self, path, *, sheet, cell, value):
        if self.fail_on == "write":
            raise RuntimeError("simulated COM failure during write")
        self.writes.append(str(path))
        data = self._cells(path)
        data.setdefault(sheet, {})
        if self.corrupt_cell is not None:
            data[sheet][self.corrupt_cell] = "CORRUPTED BY THE BACKEND"
        data[sheet][cell] = value
        Path(path).write_bytes(_OLE_HEADER + json.dumps(data).encode("utf-8"))
        return data[sheet][cell]


class WriteTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.base = Path(self.tmp.name)
        # The granted directory and a real sibling directory outside it, both inside the
        # disposable tree: nothing this suite touches may land outside self.tmp.
        self.root = self.base / "authorized"
        self.root.mkdir()
        self.outside = self.base / "outside"
        self.outside.mkdir()
        self.editor = Editor(self.root)

    def test_paths_reject_outside_and_relative(self):
        for raw in ("relative.docx", str(self.root / ".." / "outside.docx"),
                    str(self.outside / "other.docx")):
            with self.assertRaises(GuardError):
                self.editor.space.file(raw, existing=False)

    def test_paths_reject_device_namespaces_and_reserved_device_names(self):
        """Windows resolves these to devices or to a different file than the caller named."""
        real = self.root / "doc.xlsx"
        real.write_bytes(b"x")
        for raw in (f"\\\\.\\{real}", f"\\\\?\\{real}", str(real) + " ", str(real) + ".",
                    str(self.root / "CON"), str(self.root / "nul.docx"),
                    str(self.root / "COM1.xlsx"), str(self.root / "lpt9"),
                    str(self.root / "sub" / "AUX")):
            with self.subTest(raw=raw):
                with self.assertRaises(GuardError):
                    self.editor.space.file(raw)

    def test_paths_still_accept_ordinary_names_containing_reserved_words(self):
        for name in ("CONSOLE.xlsx", "NULLABLE.docx", "AUXILIARY.xlsx", "COM10.xlsx", "report.con.xlsx"):
            target = self.root / name
            target.write_bytes(b"x")
            with self.subTest(name=name):
                self.assertEqual(self.editor.space.file(str(target)), target.resolve())

    def test_paths_fail_with_one_consistent_error_type(self):
        """Every rejection is a GuardError; callers must not need to catch OSError too."""
        missing = self.root / "missing.xlsx"
        with self.assertRaisesRegex(GuardError, "could not be resolved safely"):
            self.editor.space.file(str(missing))
        # A not-yet-existing output path is the normal case and must be accepted.
        self.assertEqual(self.editor.space.file(str(missing), existing=False), missing.resolve())
        outside = self.outside / "outside-real.xlsx"
        outside.write_bytes(b"secret")
        with self.assertRaisesRegex(GuardError, "outside authorized directory"):
            self.editor.space.file(str(outside))
        (self.root / "taken.xlsx").write_bytes(b"x")
        with self.assertRaisesRegex(GuardError, "output already exists"):
            self.editor.space.file(str(self.root / "taken.xlsx"), existing=False)

    def test_word_batch_replace_needs_its_own_confirmation_and_preserves_runs(self):
        source = self.root / "batch.docx"
        doc = Document()
        first = doc.add_paragraph()
        first.add_run("Price ").bold = True          # formatting must survive the edit
        first.add_run("price and price again").italic = True
        doc.add_paragraph("Untouched paragraph mentioning nothing")
        doc.add_paragraph("Another price here")
        doc.save(source)
        original = source.read_bytes()

        request = WordBatchReplace(path=str(source), old="price", new="amount", max_occurrences=20)
        plan = self.editor.plan(request)
        self.assertTrue(plan.confirmation_required)
        self.assertIn("batch replacement", " ".join(plan.risks))
        self.assertIn("amount", plan.after)
        # Case-sensitive: "Price" in the first run is left alone, the lowercase ones change.
        self.assertEqual(plan.before.count("price"), 3)
        self.assertIn("3 occurrence(s) in 2 paragraph(s)", plan.target)

        # A bulk edit can never run without the separate confirmation.
        with self.assertRaisesRegex(GuardError, "explicit human confirmation required"):
            self.editor.execute(plan.id)
        self.assertEqual(source.read_bytes(), original)

        verified = self.editor.execute(plan.id, confirmed=True)
        output = Path(verified["output"])
        saved = Document(output)
        self.assertEqual(saved.paragraphs[0].text, "Price amount and amount again")
        self.assertEqual(saved.paragraphs[1].text, "Untouched paragraph mentioning nothing")
        self.assertEqual(saved.paragraphs[2].text, "Another amount here")
        # Run formatting is carried by the original runs, not rebuilt.
        self.assertTrue(saved.paragraphs[0].runs[0].bold)
        self.assertTrue(saved.paragraphs[0].runs[1].italic)
        self.assertEqual(source.read_bytes(), original)

        rolled_back = self.editor.rollback(verified["id"], confirmed=True)
        self.assertEqual(rolled_back["status"], "rolled_back")
        self.assertFalse(output.exists())
        self.assertEqual(source.read_bytes(), original)

    def test_word_batch_replace_refuses_too_many_occurrences_instead_of_truncating(self):
        source = self.root / "many-hits.docx"
        doc = Document()
        for index in range(10):
            doc.add_paragraph(f"price {index}")
        doc.save(source)
        original = source.read_bytes()
        with self.assertRaisesRegex(GuardError, r"found 10 occurrences, above the 4"):
            self.editor.plan(WordBatchReplace(path=str(source), old="price", new="amount",
                                              max_occurrences=4))
        # A refused bulk edit must leave nothing behind and apply nothing.
        self.assertEqual(source.read_bytes(), original)
        self.assertFalse(self.editor.audit.exists())

    def test_word_batch_replace_rejects_absent_text_and_no_op(self):
        source = self.root / "no-hits.docx"
        Document().add_paragraph("nothing here")
        doc = Document(); doc.add_paragraph("nothing here"); doc.save(source)
        with self.assertRaisesRegex(GuardError, "no occurrence"):
            self.editor.plan(WordBatchReplace(path=str(source), old="absent", new="x"))
        with self.assertRaisesRegex(GuardError, "no change"):
            self.editor.plan(WordBatchReplace(path=str(source), old="nothing", new="nothing"))

    def test_word_batch_replace_refuses_a_changed_source_between_preview_and_execute(self):
        source = self.root / "batch-changed.docx"
        doc = Document(); doc.add_paragraph("price one"); doc.add_paragraph("price two")
        doc.save(source)
        plan = self.editor.plan(WordBatchReplace(path=str(source), old="price", new="amount"))
        # A different process/user edits the file after the preview was taken.
        changed = Document(source)
        changed.paragraphs[0].text = "price one changed"
        changed.save(source)
        altered = source.read_bytes()
        with self.assertRaisesRegex(GuardError, "source changed since preview"):
            self.editor.execute(plan.id, confirmed=True)
        self.assertEqual(source.read_bytes(), altered)
        self.assertFalse(any(self.root.glob("*.agent-*")))

    def test_excel_range_trim_cleans_text_only_and_needs_confirmation(self):
        source = self.root / "clean.xlsx"
        wb = Workbook(); ws = wb.active; ws.title = "Data"
        ws["B2"] = "  padded  "
        ws["B3"] = "clean already"
        ws["B4"] = 42                      # a number is not text and must not be touched
        ws["B5"] = "=A1*3"                 # a formula is never trimmed
        ws["B6"] = "  also padded"
        ws["C2"] = "  outside the range  "
        ws["B2"].font = Font(bold=True)    # formatting must survive
        wb.save(source)
        original = source.read_bytes()

        plan = self.editor.plan(ExcelRangeTrim(path=str(source), sheet="Data",
                                               first_cell="B2", last_cell="B6"))
        self.assertTrue(plan.confirmation_required)
        self.assertIn("bulk whitespace cleaning", " ".join(plan.risks))
        self.assertIn("Data!B2:B6", plan.target)
        self.assertIn("'B2': '  padded  '", plan.before)
        self.assertIn("'B6': '  also padded'", plan.before)
        self.assertNotIn("B3", plan.after)  # already clean cells are not listed as changes
        self.assertNotIn("B4", plan.after)
        self.assertNotIn("B5", plan.after)

        with self.assertRaisesRegex(GuardError, "explicit human confirmation required"):
            self.editor.execute(plan.id)
        self.assertEqual(source.read_bytes(), original)

        verified = self.editor.execute(plan.id, confirmed=True)
        output = Path(verified["output"])
        wb2 = load_workbook(output)
        ws2 = wb2["Data"]
        self.assertEqual(ws2["B2"].value, "padded")
        self.assertEqual(ws2["B3"].value, "clean already")
        self.assertEqual(ws2["B4"].value, 42)
        self.assertEqual(ws2["B5"].value, "=A1*3")
        self.assertEqual(ws2["B6"].value, "also padded")
        self.assertEqual(ws2["C2"].value, "  outside the range  ")  # non-target untouched
        self.assertTrue(ws2["B2"].font.bold)                          # formatting survived
        wb2.close()
        self.assertEqual(source.read_bytes(), original)

        rolled_back = self.editor.rollback(verified["id"], confirmed=True)
        self.assertEqual(rolled_back["status"], "rolled_back")
        self.assertFalse(output.exists())
        self.assertEqual(source.read_bytes(), original)

    def test_excel_range_trim_refuses_oversized_ranges_and_empty_work(self):
        source = self.root / "clean-bounds.xlsx"
        wb = Workbook(); ws = wb.active; ws.title = "Data"
        ws["A1"] = "  x  "
        wb.save(source)
        with self.assertRaisesRegex(ValueError, r"bounded to 200 cells"):
            ExcelRangeTrim(path=str(source), sheet="Data", first_cell="A1", last_cell="A300")
        with self.assertRaisesRegex(ValueError, r"top-left to the bottom-right"):
            ExcelRangeTrim(path=str(source), sheet="Data", first_cell="C5", last_cell="A1")
        # Nothing to clean is refused, and the summary says how much was examined.
        with self.assertRaisesRegex(GuardError, r"no cell in B1:C2 needs whitespace cleaning \(4 cells examined\)"):
            self.editor.plan(ExcelRangeTrim(path=str(source), sheet="Data",
                                            first_cell="B1", last_cell="C2"))
        ws["A2"] = "   "; ws["B2"] = "   "   # whitespace-only text trims to empty
        wb.save(source)
        plan = self.editor.plan(ExcelRangeTrim(path=str(source), sheet="Data",
                                               first_cell="A1", last_cell="B2"))
        self.assertIn("'A2': ''", plan.after)
        self.assertFalse(self.editor.audit.exists())

    def test_excel_range_trim_refuses_a_changed_source_between_preview_and_execute(self):
        source = self.root / "clean-changed.xlsx"
        wb = Workbook(); ws = wb.active; ws.title = "Data"
        ws["A1"] = "  padded  "
        wb.save(source)
        plan = self.editor.plan(ExcelRangeTrim(path=str(source), sheet="Data",
                                               first_cell="A1", last_cell="A1"))
        wb2 = load_workbook(source); wb2["Data"]["A1"] = "  changed  "; wb2.save(source); wb2.close()
        altered = source.read_bytes()
        with self.assertRaisesRegex(GuardError, "source changed since preview"):
            self.editor.execute(plan.id, confirmed=True)
        self.assertEqual(source.read_bytes(), altered)
        self.assertFalse(any(self.root.glob("*.agent-*")))

    def test_excel_range_set_writes_a_declared_block_and_needs_its_own_confirmation(self):
        source = self.root / "block.xlsx"
        wb = Workbook(); ws = wb.active; ws.title = "Data"
        ws["A1"] = 1
        ws["A2"] = 2
        ws["C1"] = "outside the range"
        ws["C2"] = 99
        ws["A1"].font = Font(bold=True)   # formatting must survive the block write
        wb.save(source)
        original = source.read_bytes()

        plan = self.editor.plan(ExcelRangeSet(
            path=str(source), sheet="Data", first_cell="A1", last_cell="B2",
            values=[["q1", 10], [2, 5.5]], expected=[[1, None], [2, None]]))
        self.assertTrue(plan.confirmation_required)
        self.assertIn("block write", " ".join(plan.risks))
        self.assertEqual(plan.target, "Data!A1:B2")
        self.assertIn("'A1': 1", plan.before)
        self.assertIn("'B1': None", plan.before)
        self.assertIn("'A1': 'q1'", plan.after)
        self.assertIn("'B1': 10", plan.after)
        self.assertIn("'B2': 5.5", plan.after)
        # A declared cell whose value already matches is not reported as a change.
        self.assertNotIn("'A2'", plan.after)

        with self.assertRaisesRegex(GuardError, "explicit human confirmation required"):
            self.editor.execute(plan.id)
        self.assertEqual(source.read_bytes(), original)
        self.assertFalse(any(self.root.glob("*.agent-*")))

        verified = self.editor.execute(plan.id, confirmed=True)
        output = Path(verified["output"])
        wb2 = load_workbook(output)
        ws2 = wb2["Data"]
        self.assertEqual([ws2[cell].value for cell in ("A1", "B1", "A2", "B2")],
                         ["q1", 10, 2, 5.5])
        self.assertEqual([ws2[cell].value for cell in ("C1", "C2")], ["outside the range", 99])
        self.assertTrue(ws2["A1"].font.bold)  # the cell was given a new value, not a new style
        wb2.close()
        self.assertEqual(source.read_bytes(), original)
        self.assertTrue(Path(verified["backup"]).is_file())

        rolled_back = self.editor.rollback(verified["id"], confirmed=True)
        self.assertEqual(rolled_back["status"], "rolled_back")
        self.assertFalse(output.exists())
        self.assertEqual(source.read_bytes(), original)

    def test_excel_range_set_clears_a_declared_cell_and_reports_it(self):
        source = self.root / "block-clear.xlsx"
        wb = Workbook(); ws = wb.active; ws.title = "Data"
        ws["A1"] = "remove me"
        wb.save(source)
        plan = self.editor.plan(ExcelRangeSet(path=str(source), sheet="Data", first_cell="A1",
                                              last_cell="A1", values=[[None]], expected=[["remove me"]]))
        self.assertIn("'A1': 'remove me'", plan.before)
        self.assertIn("'A1': None", plan.after)
        verified = self.editor.execute(plan.id, confirmed=True)
        wb2 = load_workbook(Path(verified["output"]))
        self.assertIsNone(wb2["Data"]["A1"].value)
        wb2.close()

    def test_excel_range_set_refuses_the_whole_block_unless_every_original_matches_exactly(self):
        source = self.root / "block-mismatch.xlsx"
        wb = Workbook(); ws = wb.active; ws.title = "Data"
        ws["A1"] = 1
        ws["B1"] = "text"
        wb.save(source)
        original = source.read_bytes()
        # A wrong value, a wrong type for the same number, and a blank declared as text must
        # all refuse the edit as a whole: a shifted row can never be overwritten silently.
        for expected, values, message in (
                ([[2, "text"]], [[3, "text"]], r"A1 is 1, but expected 2"),
                ([[1.0, "text"]], [[3, "text"]], r"expected 1\.0"),
                ([[1, "other"]], [[3, "other"]], r"B1 is 'text', but expected 'other'"),
                ([[1, None]], [[3, 4]], r"B1 is 'text', but expected None")):
            with self.subTest(expected=expected):
                with self.assertRaisesRegex(GuardError, message):
                    self.editor.plan(ExcelRangeSet(path=str(source), sheet="Data", first_cell="A1",
                                                   last_cell="B1", values=values, expected=expected))
        self.assertEqual(source.read_bytes(), original)
        self.assertFalse(self.editor.audit.exists())
        self.assertFalse(any(self.root.glob("*.agent-*")))

    def test_excel_range_set_refuses_formulas_merged_cells_no_ops_and_oversized_blocks(self):
        source = self.root / "block-guards.xlsx"
        wb = Workbook(); ws = wb.active; ws.title = "Data"
        ws["A1"] = "=1+1"
        ws["B1"] = "plain"
        ws.merge_cells("C1:D1")
        ws["C1"] = "merged"
        wb.save(source)

        with self.assertRaisesRegex(GuardError, r"never overwrites\s+a formula"):
            self.editor.plan(ExcelRangeSet(path=str(source), sheet="Data", first_cell="A1",
                                           last_cell="A1", values=[["x"]],
                                           expected=[["=1+1"]]))
        with self.assertRaisesRegex(GuardError, r"C1 belongs to a merged range|D1 belongs to a merged range"):
            self.editor.plan(ExcelRangeSet(path=str(source), sheet="Data", first_cell="C1",
                                           last_cell="D1", values=[["x", "y"]], expected=[["merged", None]]))
        with self.assertRaisesRegex(GuardError, "would make no change"):
            self.editor.plan(ExcelRangeSet(path=str(source), sheet="Data", first_cell="B1",
                                           last_cell="B1", values=[["plain"]], expected=[["plain"]]))
        with self.assertRaisesRegex(GuardError, "sheet not found"):
            self.editor.plan(ExcelRangeSet(path=str(source), sheet="Missing", first_cell="B1",
                                           last_cell="B1", values=[["x"]], expected=[["plain"]]))
        # A model-supplied formula, a wrong-shaped block and an oversized range are rejected
        # before any file is opened.
        with self.assertRaisesRegex(ValueError, "literal values only"):
            ExcelRangeSet(path=str(source), sheet="Data", first_cell="B1", last_cell="B1",
                          values=[["=SUM(A1:A2)"]], expected=[["plain"]])
        with self.assertRaisesRegex(ValueError, r"must be a 2x1 block"):
            ExcelRangeSet(path=str(source), sheet="Data", first_cell="B1", last_cell="B2",
                          values=[["only one row"]], expected=[["plain"], [None]])
        with self.assertRaisesRegex(ValueError, r"2x1 block"):
            ExcelRangeSet(path=str(source), sheet="Data", first_cell="B1", last_cell="B2",
                          values=[["a"], ["b"]], expected=[["plain"]])
        with self.assertRaisesRegex(ValueError, r"write 6 instead of 6\.0"):
            # openpyxl normalises a whole float to an integer, so the write could not verify.
            ExcelRangeSet(path=str(source), sheet="Data", first_cell="B1", last_cell="B1",
                          values=[[6.0]], expected=[["plain"]])
        with self.assertRaisesRegex(ValueError, "bounded to 100 cells"):
            ExcelRangeSet(path=str(source), sheet="Data", first_cell="A1", last_cell="A101",
                          values=[[1]] * 101, expected=[[None]] * 101)
        self.assertFalse(self.editor.audit.exists())

    def test_excel_range_set_refuses_a_changed_source_between_preview_and_execute(self):
        source = self.root / "block-changed.xlsx"
        wb = Workbook(); ws = wb.active; ws.title = "Data"
        ws["A1"] = "original"
        wb.save(source)
        plan = self.editor.plan(ExcelRangeSet(path=str(source), sheet="Data", first_cell="A1",
                                              last_cell="A1", values=[["new"]], expected=[["original"]]))
        wb2 = load_workbook(source); wb2["Data"]["A1"] = "changed"; wb2.save(source); wb2.close()
        altered = source.read_bytes()
        with self.assertRaisesRegex(GuardError, "source changed since preview"):
            self.editor.execute(plan.id, confirmed=True)
        self.assertEqual(source.read_bytes(), altered)
        self.assertFalse(any(self.root.glob("*.agent-*")))
        self.assertFalse(self.editor.audit.exists())

    def test_excel_range_set_refuses_a_stale_preview_instead_of_writing_it(self):
        """The declared originals are revalidated at execute time, not trusted from the plan."""
        from dataclasses import replace

        source = self.root / "block-stale.xlsx"
        wb = Workbook(); ws = wb.active; ws.title = "Data"
        ws["A1"] = "original"
        wb.save(source)
        original = source.read_bytes()
        plan = self.editor.plan(ExcelRangeSet(path=str(source), sheet="Data", first_cell="A1",
                                             last_cell="A1", values=[["new"]], expected=[["original"]]))
        # Simulate a preview that no longer describes the file the execution would edit.
        self.editor.plans[plan.id] = replace(plan, after="{'A1': 'stale'}")
        with self.assertRaisesRegex(GuardError, "block write target changed since preview"):
            self.editor.execute(plan.id, confirmed=True)
        self.assertEqual(source.read_bytes(), original)
        self.assertFalse(any(self.root.glob("*.agent-*")))

    def test_word_table_row_insert_and_delete_with_verification(self):
        source = self.root / "rows.docx"
        doc = Document()
        doc.add_paragraph("Intro stays put")
        table = doc.add_table(rows=3, cols=2)
        for index, row in enumerate(table.rows):
            row.cells[0].text = f"label {index}"
            row.cells[1].text = f"value {index}"
        doc.add_paragraph("Outro stays put")
        doc.save(source)
        original = source.read_bytes()

        plan = self.editor.plan(WordTableRow(path=str(source), table_index=0, row_index=0,
                                             action="insert_after"))
        self.assertEqual(plan.target, "table[0].row[0] insert_after")
        self.assertEqual(plan.before, plan.before.replace(plan.before, plan.before))
        self.assertIn("'rows': 3", plan.before)
        self.assertIn("'rows': 4", plan.after)
        self.assertFalse(plan.confirmation_required)  # inserting an empty row is not destructive

        verified = self.editor.execute(plan.id, confirmed=True)
        output = Path(verified["output"])
        saved = Document(output)
        self.assertEqual(len(saved.tables[0].rows), 4)
        self.assertEqual([c.text for c in saved.tables[0].rows[0].cells], ["label 0", "value 0"])
        self.assertEqual([c.text for c in saved.tables[0].rows[1].cells], ["", ""])  # the new row
        self.assertEqual([c.text for c in saved.tables[0].rows[2].cells], ["label 1", "value 1"])
        self.assertEqual(saved.paragraphs[0].text, "Intro stays put")
        self.assertEqual(saved.paragraphs[-1].text, "Outro stays put")
        self.assertEqual(source.read_bytes(), original)

        # Deleting a row is destructive and therefore needs its own confirmation.
        delete_plan = self.editor.plan(WordTableRow(path=str(source), table_index=0,
                                                    row_index=1, action="delete"))
        self.assertTrue(delete_plan.confirmation_required)
        self.assertIn("deletes a table row", " ".join(delete_plan.risks))
        self.assertIn("'deleted_row_content': ['label 1', 'value 1']", delete_plan.before)
        # The post-edit summary must not reuse the same index in the shortened table and
        # report a different row as the deleted one.
        self.assertEqual(delete_plan.after,
                         "{'table_index': 0, 'rows': 2, 'columns': 2}")
        with self.assertRaisesRegex(GuardError, "explicit human confirmation required"):
            self.editor.execute(delete_plan.id)
        self.assertEqual(source.read_bytes(), original)

        removed = self.editor.execute(delete_plan.id, confirmed=True)
        saved = Document(Path(removed["output"]))
        self.assertEqual(len(saved.tables[0].rows), 2)
        self.assertEqual([[c.text for c in row.cells] for row in saved.tables[0].rows],
                         [["label 0", "value 0"], ["label 2", "value 2"]])
        self.assertEqual(source.read_bytes(), original)

    def test_word_table_row_insert_before_and_at_rejects_merged_or_nested_tables(self):
        source = self.root / "merged.docx"
        doc = Document()
        table = doc.add_table(rows=2, cols=3)
        table.cell(0, 0).merge(table.cell(0, 1))          # horizontal merge in row 0
        doc.save(source)
        # A merge anywhere makes the grid ambiguous to restructure, so the whole table is
        # refused rather than only the merged rows.
        for row, action in ((0, "insert_after"), (1, "insert_before"), (1, "insert_after"),
                            (0, "delete"), (1, "delete")):
            with self.subTest(row=row, action=action):
                with self.assertRaisesRegex(GuardError, "merged cells"):
                    self.editor.plan(WordTableRow(path=str(source), table_index=0,
                                                  row_index=row, action=action))
        self.assertFalse(self.editor.audit.exists())

    def test_word_table_row_insert_before_on_a_plain_table(self):
        source = self.root / "plain.docx"
        doc = Document()
        table = doc.add_table(rows=2, cols=3)
        for index, row in enumerate(table.rows):
            for column, cell in enumerate(row.cells):
                cell.text = f"r{index}c{column}"
        doc.save(source)
        plan = self.editor.plan(WordTableRow(path=str(source), table_index=0, row_index=1,
                                             action="insert_before"))
        self.assertEqual(plan.target, "table[0].row[1] insert_before")
        verified = self.editor.execute(plan.id, confirmed=True)
        saved = Document(Path(verified["output"]))
        self.assertEqual([[c.text for c in row.cells] for row in saved.tables[0].rows],
                         [["r0c0", "r0c1", "r0c2"], ["", "", ""], ["r1c0", "r1c1", "r1c2"]])

    def test_word_table_row_refuses_out_of_bounds_and_the_last_remaining_row(self):
        source = self.root / "tiny.docx"
        doc = Document()
        doc.add_table(rows=1, cols=2)
        doc.save(source)
        with self.assertRaisesRegex(GuardError, "only row"):
            self.editor.plan(WordTableRow(path=str(source), table_index=0, row_index=0,
                                          action="delete"))
        with self.assertRaisesRegex(GuardError, "row index 1 is out of bounds"):
            self.editor.plan(WordTableRow(path=str(source), table_index=0, row_index=1,
                                          action="insert_after"))
        with self.assertRaisesRegex(GuardError, "table index 2 is out of bounds"):
            self.editor.plan(WordTableRow(path=str(source), table_index=2, row_index=0,
                                          action="insert_after"))
        with self.assertRaisesRegex(GuardError, "nested table"):
            nested = self.root / "nested.docx"
            outer = Document()
            inner = outer.add_table(rows=1, cols=1).cell(0, 0).add_table(rows=1, cols=1)
            inner.cell(0, 0).text = "inner"
            outer.save(nested)
            self.editor.plan(WordTableRow(path=str(nested), table_index=0, row_index=0,
                                          action="insert_after"))

    def _merged_source(self, name="merged-cell.docx"):
        """3x3 table with a gridSpan pair (row 0), a vMerge pair (column 0) and styled text."""
        source = self.root / name
        doc = Document()
        table = doc.add_table(rows=3, cols=3)
        table.style = "Table Grid"
        anchor = table.cell(0, 0)
        anchor.text = "span A"
        anchor.paragraphs[0].runs[0].bold = True          # formatting must survive the edit
        anchor.merge(table.cell(0, 1))
        table.cell(0, 2).text = "span B"
        table.cell(1, 0).text = "vertical"
        table.cell(1, 0).merge(table.cell(2, 0))
        table.cell(1, 1).text = "plain 1"
        table.cell(2, 1).text = "plain 2"
        doc.save(source)
        return source

    def test_word_merged_cell_replace_keeps_the_merge_and_needs_confirmation(self):
        source = self._merged_source()
        original = source.read_bytes()
        plan = self.editor.plan(WordMergedCellText(
            path=str(source), table_index=0, row_index=0, column_index=0,
            action="replace", old="span A", new="span A v2"))
        self.assertTrue(plan.confirmation_required)
        self.assertIn("merged-cell text edit", " ".join(plan.risks))
        self.assertIn("gridSpan=2", plan.target)          # the preview names the merge it writes into
        self.assertEqual(plan.before, "span A")
        self.assertEqual(plan.after, "span A v2")

        with self.assertRaisesRegex(GuardError, "explicit human confirmation required"):
            self.editor.execute(plan.id)
        self.assertEqual(source.read_bytes(), original)

        verified = self.editor.execute(plan.id, confirmed=True)
        output = Path(verified["output"])
        saved = Document(output)
        self.assertEqual(saved.tables[0].cell(0, 0).text, "span A v2")
        self.assertTrue(saved.tables[0].cell(0, 0).paragraphs[0].runs[0].bold)
        # The merge is byte-for-byte the same shape it was before the write.
        self.assertEqual(_merge_projection(saved), _merge_projection(Document(source)))
        self.assertEqual(saved.tables[0].cell(0, 1).text, "span A v2")  # the span still reads through
        self.assertEqual(saved.tables[0].cell(0, 2).text, "span B")
        self.assertEqual(source.read_bytes(), original)

        rolled_back = self.editor.rollback(verified["id"], confirmed=True)
        self.assertEqual(rolled_back["status"], "rolled_back")
        self.assertFalse(output.exists())
        self.assertEqual(source.read_bytes(), original)

    def test_word_merged_cell_append_adds_one_paragraph_and_keeps_the_merge(self):
        source = self._merged_source("merged-append.docx")
        original = source.read_bytes()
        plan = self.editor.plan(WordMergedCellText(
            path=str(source), table_index=0, row_index=1, column_index=0,
            action="append", new="appended line"))
        self.assertTrue(plan.confirmation_required)
        self.assertEqual(plan.before, "vertical")
        self.assertEqual(plan.after, "vertical\nappended line")
        self.assertIn("vMerge=restart", plan.target)

        verified = self.editor.execute(plan.id, confirmed=True)
        output = Path(verified["output"])
        saved = Document(output)
        cell = saved.tables[0].cell(1, 0)
        self.assertEqual(cell.text, "vertical\nappended line")
        # The continuation cell still carries no text, read through the raw-XML structure
        # report rather than python-docx's collapsed cell mapping.
        structure = read_word_table_structure(output, limit=10)
        continuation = structure["rows"][2]["cells"][0]
        self.assertEqual(continuation["vmerge"], "continue")
        self.assertEqual(continuation["text"], "")
        self.assertTrue(structure["merge_consistent"])
        self.assertEqual(_merge_projection(saved), _merge_projection(Document(source)))
        self.assertEqual(source.read_bytes(), original)

    def test_word_merged_cell_write_refuses_covered_and_continued_coordinates(self):
        source = self._merged_source("merged-refuse.docx")
        original = source.read_bytes()
        # A gridSpan follower and a vMerge continuation do not carry text: both are refused
        # and the refusal names the origin cell to write through instead.
        with self.assertRaisesRegex(GuardError, r"跨列合并覆盖"):
            self.editor.plan(WordMergedCellText(path=str(source), table_index=0, row_index=0,
                                                column_index=1, action="replace",
                                                old="span A", new="x"))
        with self.assertRaisesRegex(GuardError, r"文本由 \(行 1, 列 0\) 承载"):
            self.editor.plan(WordMergedCellText(path=str(source), table_index=0, row_index=2,
                                                column_index=0, action="append", new="x"))
        with self.assertRaisesRegex(GuardError, "超出表格范围"):
            self.editor.plan(WordMergedCellText(path=str(source), table_index=0, row_index=7,
                                                column_index=0, action="append", new="x"))
        with self.assertRaisesRegex(GuardError, "table index 1 is out of bounds"):
            self.editor.plan(WordMergedCellText(path=str(source), table_index=1, row_index=0,
                                                column_index=0, action="append", new="x"))
        with self.assertRaisesRegex(GuardError, "期望在目标单元格段落中找到 2 处"):
            self.editor.plan(WordMergedCellText(path=str(source), table_index=0, row_index=0,
                                                column_index=0, action="replace",
                                                old="span A", new="x", expected_count=2))
        self.assertEqual(source.read_bytes(), original)
        self.assertFalse(self.editor.audit.exists())

    def test_word_merged_cell_write_refuses_a_structurally_broken_table(self):
        source = self.root / "merged-broken.docx"
        doc = Document()
        table = doc.add_table(rows=2, cols=2)
        table.cell(0, 0).text = "a"
        table.cell(0, 1).text = "b"
        table.cell(1, 0).text = "c"
        table.cell(1, 1).text = "d"
        doc.save(source)
        document = Document(source)
        row = document.tables[0]._tbl.findall(qn("w:tr"))[1]
        row.remove(row.findall(qn("w:tc"))[-1])       # the row no longer covers the grid
        document.save(source)
        with self.assertRaisesRegex(GuardError, "表格存在结构不一致"):
            self.editor.plan(WordMergedCellText(path=str(source), table_index=0, row_index=0,
                                                column_index=0, action="append", new="x"))

    def test_merge_projection_makes_a_broken_merge_visible(self):
        """The write check must have teeth: reshaping a merge has to change the projection."""
        source = self._merged_source("merged-projection.docx")
        doc = Document(source)
        before = _merge_projection(doc)
        self.assertEqual(before, _merge_projection(Document(source)))
        tbl = doc.tables[0]._tbl
        span = tbl.findall(qn("w:tr"))[0].findall(qn("w:tc"))[0].find(qn("w:tcPr")).find(qn("w:gridSpan"))
        span.set(qn("w:val"), "1")                     # pretend a write broke the horizontal merge
        self.assertNotEqual(_merge_projection(doc), before)
        properties = tbl.findall(qn("w:tr"))[2].findall(qn("w:tc"))[0].find(qn("w:tcPr"))
        merge = properties.find(qn("w:vMerge"))
        self.assertIsNotNone(merge)
        properties.remove(merge)                       # pretend a write dropped the continuation
        self.assertNotEqual(_merge_projection(doc), before)

    def test_word_merged_cell_write_refuses_a_changed_source(self):
        source = self._merged_source("merged-changed.docx")
        plan = self.editor.plan(WordMergedCellText(path=str(source), table_index=0, row_index=0,
                                                   column_index=0, action="replace",
                                                   old="span A", new="span A v2"))
        document = Document(source)
        document.tables[0].cell(0, 2).text = "changed elsewhere"
        document.save(source)
        altered = source.read_bytes()
        with self.assertRaisesRegex(GuardError, "source changed since preview"):
            self.editor.execute(plan.id, confirmed=True)
        self.assertEqual(source.read_bytes(), altered)
        self.assertFalse(any(self.root.glob("*.agent-*")))

    def _revision_source(self, name="revisions.docx"):
        """One inserted run, one deleted run and one formatting revision, all in the body."""
        source = self.root / name
        doc = Document()
        doc.sections[0].header.paragraphs[0].text = "header"   # materialises word/header1.xml
        inserted = doc.add_paragraph("keep ")
        insertion = OxmlElement("w:ins")
        insertion.set(qn("w:id"), "1")
        insertion.set(qn("w:author"), "Alice")
        run = OxmlElement("w:r")
        text = OxmlElement("w:t")
        text.text = "inserted"
        run.append(text)
        insertion.append(run)
        inserted._p.append(insertion)
        inserted.add_run(" end")

        deleted = doc.add_paragraph("keep ")
        deletion = OxmlElement("w:del")
        deletion.set(qn("w:id"), "2")
        deletion.set(qn("w:author"), "Bob")
        run = OxmlElement("w:r")
        deleted_text = OxmlElement("w:delText")
        deleted_text.text = "removed"
        run.append(deleted_text)
        deletion.append(run)
        deleted._p.append(deletion)
        deleted.add_run(" end")

        styled = doc.add_paragraph()
        styled_run = styled.add_run("styled")
        styled_run.bold = True
        change = OxmlElement("w:rPrChange")
        change.set(qn("w:id"), "3")
        change.set(qn("w:author"), "Alice")
        change.append(OxmlElement("w:rPr"))        # the previous state: not bold
        styled_run._r.get_or_add_rPr().append(change)
        doc.save(source)
        return source

    def _revision_mark_count(self, path):
        with zipfile.ZipFile(path) as archive:
            document = archive.read("word/document.xml").decode("utf-8")
        return sum(document.count(f"<w:{name}") for name in
                   ("ins", "del", "rPrChange", "pPrChange"))

    def test_accept_all_revisions_keeps_insertions_and_drops_deleted_content(self):
        source = self._revision_source()
        original = source.read_bytes()
        plan = self.editor.plan(WordRevisionDecision(path=str(source), action="accept_all"))
        self.assertTrue(plan.confirmation_required)
        self.assertIn("deciding all revisions", " ".join(plan.risks))
        self.assertIn("接受全部修订", plan.target)
        preview = json.loads(plan.after)
        self.assertEqual([item["after"] for item in preview["changed"]], ["keep inserted end"])

        with self.assertRaisesRegex(GuardError, "explicit human confirmation required"):
            self.editor.execute(plan.id)
        self.assertEqual(source.read_bytes(), original)

        verified = self.editor.execute(plan.id, confirmed=True)
        output = Path(verified["output"])
        saved = Document(output)
        # 接受 = w:ins 解包保留内容；w:del 连同被删内容彻底消失。
        self.assertEqual([p.text for p in saved.paragraphs], ["keep inserted end", "keep  end", "styled"])
        self.assertTrue(saved.paragraphs[2].runs[0].bold)      # current formatting is kept
        self.assertEqual(self._revision_mark_count(output), 0)
        self.assertEqual(source.read_bytes(), original)

        rolled_back = self.editor.rollback(verified["id"], confirmed=True)
        self.assertEqual(rolled_back["status"], "rolled_back")
        self.assertFalse(output.exists())
        self.assertEqual(source.read_bytes(), original)

    def test_reject_all_revisions_drops_insertions_and_restores_deleted_text(self):
        source = self._revision_source("revisions-reject.docx")
        original = source.read_bytes()
        plan = self.editor.plan(WordRevisionDecision(path=str(source), action="reject_all"))
        self.assertTrue(plan.confirmation_required)
        preview = json.loads(plan.after)
        self.assertEqual([item["after"] for item in preview["changed"]], ["keep removed end"])

        with self.assertRaisesRegex(GuardError, "explicit human confirmation required"):
            self.editor.execute(plan.id)
        self.assertEqual(source.read_bytes(), original)

        verified = self.editor.execute(plan.id, confirmed=True)
        output = Path(verified["output"])
        saved = Document(output)
        # 拒绝 = w:ins 连同插入内容移除；w:del 解包且 w:delText 还原为 w:t。
        self.assertEqual([p.text for p in saved.paragraphs],
                         ["keep  end", "keep removed end", "styled"])
        self.assertIsNone(saved.paragraphs[2].runs[0].bold)    # previous formatting is restored
        self.assertEqual(self._revision_mark_count(output), 0)
        self.assertEqual(source.read_bytes(), original)

    def test_revision_decision_refuses_marks_it_cannot_decide_without_a_half_result(self):
        cases = {}
        doc = Document()
        paragraph = doc.add_paragraph("moved text")
        paragraph._p.append(OxmlElement("w:moveFrom"))          # a move, not a plain revision
        cases["moveFrom"] = doc
        doc = Document()
        table = doc.add_table(rows=1, cols=1)
        table.cell(0, 0).text = "cell"
        row_properties = OxmlElement("w:trPr")
        row_insert = OxmlElement("w:ins")
        row_insert.set(qn("w:id"), "9")
        row_properties.append(row_insert)
        table.rows[0]._tr.insert(0, row_properties)             # a row-level insertion
        cases["ins@trPr"] = doc
        doc = Document()
        paragraph = doc.add_paragraph("paragraph mark")
        properties = OxmlElement("w:pPr")
        run_properties = OxmlElement("w:rPr")
        run_properties.append(OxmlElement("w:del"))             # deleted paragraph mark
        properties.append(run_properties)
        paragraph._p.insert(0, properties)
        cases["del@rPr"] = doc

        for expected, document in cases.items():
            with self.subTest(mark=expected):
                source = self.root / f"unsupported-{expected.replace('@', '-')}.docx"
                document.save(source)
                original = source.read_bytes()
                with self.assertRaisesRegex(GuardError, "本操作不支持的修订标记"):
                    self.editor.plan(WordRevisionDecision(path=str(source), action="accept_all"))
                with self.assertRaisesRegex(GuardError, "本操作不支持的修订标记"):
                    self.editor.plan(WordRevisionDecision(path=str(source), action="reject_all"))
                self.assertEqual(source.read_bytes(), original)
        self.assertFalse(self.editor.audit.exists())

    def test_revision_decision_refuses_revisions_outside_the_body_part(self):
        source = self._revision_source("header-revisions.docx")
        with zipfile.ZipFile(source) as archive:
            parts = {name: archive.read(name) for name in archive.namelist()}
        header = parts["word/header1.xml"].decode("utf-8")
        parts["word/header1.xml"] = header.replace(
            "</w:hdr>",
            '<w:ins w:id="42" w:author="Alice"><w:r><w:t>header addition</w:t></w:r></w:ins></w:hdr>'
        ).encode("utf-8")
        with zipfile.ZipFile(source, "w", zipfile.ZIP_DEFLATED) as archive:
            for name, data in parts.items():
                archive.writestr(name, data)

        with self.assertRaisesRegex(GuardError, "正文之外的部件中仍有修订标记"):
            self.editor.plan(WordRevisionDecision(path=str(source), action="accept_all"))
        self.assertFalse(self.editor.audit.exists())

    def test_revision_decision_refuses_a_document_without_revisions(self):
        source = self.root / "no-revisions.docx"
        doc = Document()
        doc.add_paragraph("nothing to decide")
        doc.save(source)
        with self.assertRaisesRegex(GuardError, "没有可处理的修订标记"):
            self.editor.plan(WordRevisionDecision(path=str(source), action="reject_all"))

    def test_revision_decision_still_refuses_other_advanced_xml(self):
        """The tolerance is narrow: a content control alongside revisions keeps it refused."""
        source = self._revision_source("revisions-with-control.docx")
        document = Document(source)
        structured = OxmlElement("w:sdt")
        properties = OxmlElement("w:sdtPr")
        tag = OxmlElement("w:tag")
        tag.set(qn("w:val"), "Field")
        properties.append(tag)
        structured.append(properties)
        content = OxmlElement("w:sdtContent")
        paragraph = OxmlElement("w:p")
        run = OxmlElement("w:r")
        text = OxmlElement("w:t")
        text.text = "control"
        run.append(text)
        paragraph.append(run)
        content.append(paragraph)
        structured.append(content)
        document.element.body.append(structured)
        document.save(source)

        with self.assertRaisesRegex(GuardError, "unsupported/risky structures"):
            self.editor.plan(WordRevisionDecision(path=str(source), action="accept_all"))

    def test_revision_decision_refuses_a_changed_source(self):
        source = self._revision_source("revisions-changed.docx")
        plan = self.editor.plan(WordRevisionDecision(path=str(source), action="accept_all"))
        document = Document(source)
        document.add_paragraph("added after the preview")
        document.save(source)
        altered = source.read_bytes()
        with self.assertRaisesRegex(GuardError, "source changed since preview"):
            self.editor.execute(plan.id, confirmed=True)
        self.assertEqual(source.read_bytes(), altered)
        self.assertFalse(any(self.root.glob("*.agent-*")))

    def _legacy_doc(self, name="legacy.doc", paragraphs=("First paragraph", "price is 5",
                                                         "Last paragraph")):
        path = self.root / name
        path.write_bytes(_OLE_HEADER + "\n".join(paragraphs).encode("utf-8"))
        return path

    def _legacy_xls(self, name="legacy.xls", cells=None):
        path = self.root / name
        path.write_bytes(_OLE_HEADER + json.dumps({"Data": cells or {"A1": 4, "B1": "keep"}}).encode("utf-8"))
        return path

    def test_legacy_doc_replace_chain_runs_on_the_audit_copy_only(self):
        """Full chain with an injected backend: plan -> confirm -> copy -> verify -> rollback."""
        source = self._legacy_doc()
        original = source.read_bytes()
        backend = FakeLegacyBackend()
        editor = Editor(self.root, legacy_backend=backend)

        plan = editor.plan(LegacyWordReplace(path=str(source), paragraph_index=1,
                                             old="price", new="amount"))
        self.assertTrue(plan.confirmation_required)
        self.assertIn("legacy .doc/.xls edit through Office COM", " ".join(plan.risks))
        self.assertIn("paragraph[1]", plan.target)
        self.assertEqual(plan.before, "price is 5")
        self.assertEqual(plan.after, "amount is 5")

        with self.assertRaisesRegex(GuardError, "explicit human confirmation required"):
            editor.execute(plan.id)
        self.assertEqual(source.read_bytes(), original)
        self.assertFalse(any(self.root.glob("*.agent-*")))
        self.assertEqual(backend.probes, ["word"])                 # probed once, then cached

        verified = editor.execute(plan.id, confirmed=True)
        output = Path(verified["output"])
        self.assertEqual(output.suffix, ".doc")
        self.assertEqual(backend.read_word_paragraphs(str(output), limit=5),
                         ["First paragraph", "amount is 5", "Last paragraph"])
        self.assertEqual(source.read_bytes(), original)
        # Every COM write went to the audit job's own copy, never to the user's file.
        self.assertTrue(backend.writes)
        self.assertNotIn(str(source), backend.writes)
        self.assertTrue(all(str(self.editor.audit) in path for path in backend.writes))
        self.assertEqual(verified["legacy_verification"]["kind"], "paragraph_window")

        rolled_back = editor.rollback(verified["id"], confirmed=True)
        self.assertEqual(rolled_back["status"], "rolled_back")
        self.assertFalse(output.exists())
        self.assertEqual(source.read_bytes(), original)

    def test_legacy_cell_set_chain_and_exact_original_match(self):
        source = self._legacy_xls()
        original = source.read_bytes()
        backend = FakeLegacyBackend()
        editor = Editor(self.root, legacy_backend=backend)

        plan = editor.plan(LegacyCellSet(path=str(source), sheet="Data", cell="A1", value=7,
                                         expected_original=4))
        self.assertTrue(plan.confirmation_required)
        self.assertEqual((plan.before, plan.after), ("4", "7"))
        with self.assertRaisesRegex(GuardError, "explicit human confirmation required"):
            editor.execute(plan.id)
        self.assertEqual(source.read_bytes(), original)

        verified = editor.execute(plan.id, confirmed=True)
        output = Path(verified["output"])
        self.assertEqual(backend.read_excel_region(str(output), sheet="Data", first_cell="A1",
                                                   last_cell="T20"), {"A1": 7, "B1": "keep"})
        self.assertEqual(source.read_bytes(), original)
        self.assertEqual(verified["legacy_verification"]["window"], "A1:T20")
        self.assertNotIn(str(source), backend.writes)

    def test_legacy_plan_refuses_mismatched_or_pointless_inputs(self):
        source = self._legacy_xls("refuse.xls")
        doc = self._legacy_doc("refuse.doc")
        backend = FakeLegacyBackend()
        editor = Editor(self.root, legacy_backend=backend)
        with self.assertRaisesRegex(GuardError, "does not match Data!A1"):
            editor.plan(LegacyCellSet(path=str(source), sheet="Data", cell="A1", value=7,
                                      expected_original=99))
        with self.assertRaisesRegex(GuardError, "would make no change"):
            editor.plan(LegacyCellSet(path=str(source), sheet="Data", cell="A1", value=4,
                                      expected_original=4))
        with self.assertRaisesRegex(GuardError, r"expected 2 occurrence\(s\) of the literal, found 1"):
            editor.plan(LegacyWordReplace(path=str(doc), paragraph_index=1, old="price",
                                          new="amount", expected_count=2))
        with self.assertRaisesRegex(GuardError, "out of bounds"):
            editor.plan(LegacyWordReplace(path=str(doc), paragraph_index=9, old="price",
                                          new="amount"))
        with self.assertRaisesRegex(GuardError, r"\.xls required"):
            editor.plan(LegacyCellSet(path=str(doc), sheet="Data", cell="A1", value=1,
                                      expected_original=None))
        with self.assertRaisesRegex(GuardError, r"\.doc required"):
            editor.plan(LegacyWordReplace(path=str(source), old="a", new="b"))
        with self.assertRaisesRegex(ValueError, "literal values only"):
            LegacyCellSet(path=str(source), sheet="Data", cell="A1", value="=1+1",
                          expected_original=4)
        with self.assertRaisesRegex(ValueError, "would make no change"):
            LegacyWordReplace(path=str(doc), old="same", new="same")
        # Every refusal above happened before any write: both sources are untouched.
        self.assertEqual(backend.writes, [])
        self.assertFalse(self.editor.audit.exists())
        self.assertFalse(any(self.root.glob("*.agent-*")))
        self.assertEqual(source.read_bytes(), _OLE_HEADER + json.dumps(
            {"Data": {"A1": 4, "B1": "keep"}}).encode("utf-8"))
        self.assertEqual(doc.read_bytes(), _OLE_HEADER + "First paragraph\nprice is 5\nLast paragraph".encode("utf-8"))

    def test_legacy_source_must_be_a_real_ole_file(self):
        fake = self.root / "not-really.doc"
        fake.write_text("this is plain text, not an OLE compound file", encoding="utf-8")
        editor = Editor(self.root, legacy_backend=FakeLegacyBackend())
        with self.assertRaisesRegex(GuardError, "OLE compound-file header"):
            editor.plan(LegacyWordReplace(path=str(fake), old="plain", new="x"))
        # Classification trusts the extension, so the header check is what actually stops a
        # renamed OOXML or plain file from entering the COM channel.
        renamed = self.root / "renamed.doc"
        document = Document()
        document.add_paragraph("price")
        document.save(renamed)
        with self.assertRaisesRegex(GuardError, "OLE compound-file header"):
            editor.plan(LegacyWordReplace(path=str(renamed), old="price", new="x"))
        self.assertFalse(editor.audit.exists())
        self.assertFalse(any(self.root.glob("*.agent-*")))

    def test_legacy_channel_stays_closed_without_a_backend_or_a_passing_probe(self):
        source = self._legacy_doc("closed.doc")
        original = source.read_bytes()
        with self.assertRaisesRegex(GuardError, "未注入 Office COM 后端"):
            Editor(self.root).plan(LegacyWordReplace(path=str(source), old="price", new="amount"))

        backend = FakeLegacyBackend(available=False, reason="WinError 5: gen_py denied")
        editor = Editor(self.root, legacy_backend=backend)
        for _ in range(2):                                         # the probe result is cached
            with self.assertRaisesRegex(GuardError, "未检测到可用的 Microsoft Word COM 写通道"):
                editor.plan(LegacyWordReplace(path=str(source), old="price", new="amount"))
        self.assertIn("WinError 5: gen_py denied", str(
            editor._legacy_probes["word"]["reason"]))
        self.assertEqual(backend.probes, ["word"])
        self.assertEqual(backend.writes, [])
        self.assertEqual(source.read_bytes(), original)
        self.assertFalse(editor.audit.exists())                    # a refusal leaves no trace

    def test_legacy_write_failure_keeps_the_source_and_audits_it_as_failed(self):
        source = self._legacy_doc("failing.doc")
        original = source.read_bytes()
        backend = FakeLegacyBackend(fail_on="write")
        editor = Editor(self.root, legacy_backend=backend)
        plan = editor.plan(LegacyWordReplace(path=str(source), paragraph_index=2,
                                             old="Last", new="Final"))
        with self.assertRaisesRegex(RuntimeError, "simulated COM failure"):
            editor.execute(plan.id, confirmed=True)
        self.assertEqual(source.read_bytes(), original)
        self.assertFalse(any(self.root.glob("*.agent-*")))
        record = json.loads((editor.audit / plan.id / "audit.json").read_text(encoding="utf-8"))
        self.assertEqual(record["status"], "failed")               # never audited as success
        self.assertFalse((editor.audit / plan.id / "working.doc").exists())

    def test_legacy_verification_catches_a_changed_neighbour(self):
        """The reopened-copy check must notice anything outside the edited cell/paragraph."""
        source = self._legacy_doc("neighbour.doc")
        original = source.read_bytes()
        editor = Editor(self.root, legacy_backend=FakeLegacyBackend(corrupt_paragraph=0))
        plan = editor.plan(LegacyWordReplace(path=str(source), paragraph_index=1,
                                             old="price", new="amount"))
        with self.assertRaisesRegex(GuardError, "legacy 非目标段落发生变化"):
            editor.execute(plan.id, confirmed=True)
        self.assertEqual(source.read_bytes(), original)
        self.assertFalse(any(self.root.glob("*.agent-*")))

        cells = self._legacy_xls("neighbour.xls")
        cells_before = cells.read_bytes()
        editor = Editor(self.root, legacy_backend=FakeLegacyBackend(corrupt_cell="B1"))
        plan = editor.plan(LegacyCellSet(path=str(cells), sheet="Data", cell="A1", value=7,
                                         expected_original=4))
        with self.assertRaisesRegex(GuardError, "legacy 非目标单元格发生变化"):
            editor.execute(plan.id, confirmed=True)
        self.assertEqual(cells.read_bytes(), cells_before)

    def test_legacy_channel_can_only_write_inside_the_audit_job(self):
        source = self._legacy_doc("target-guard.doc")
        editor = Editor(self.root, legacy_backend=FakeLegacyBackend())
        for candidate in (source, Path(str(source) + ".copy"), self.root / "elsewhere.doc"):
            with self.subTest(candidate=str(candidate)):
                with self.assertRaisesRegex(GuardError, "audit job's own copy"):
                    editor._legacy_working_copy(candidate)
        inside = editor.audit / "job" / "working.doc"
        self.assertEqual(editor._legacy_working_copy(inside), inside)

    def test_legacy_accepts_only_the_matching_legacy_extension(self):
        excel_source = self._legacy_xls("typed.xls")
        editor = Editor(self.root, legacy_backend=FakeLegacyBackend())
        with self.assertRaisesRegex(GuardError, r"\.doc required"):
            editor.plan(LegacyWordReplace(path=str(excel_source), old="a", new="b"))

    def _merge_source(self, name, texts, heading=None, table=False):
        path = self.root / name
        document = Document()
        if heading is not None:
            document.add_heading(heading, level=1)
        for text in texts:
            document.add_paragraph(text)
        if table:
            cells = document.add_table(rows=1, cols=2)
            cells.cell(0, 0).text = f"{name}-a"
            cells.cell(0, 1).text = f"{name}-b"
        document.save(path)
        return path

    def test_docx_merge_creates_one_new_document_and_keeps_every_source(self):
        first = self._merge_source("one.docx", ["alpha one"], "First", table=True)
        second = self._merge_source("two.docx", ["beta two"], "Second")
        third = self._merge_source("three.docx", ["gamma three"])
        digests = {path.name: digest(path) for path in (first, second, third)}
        plan = self.editor.plan(DocxMerge(path=str(first),
                                          sources=["one.docx", "two.docx", "three.docx"]))
        self.assertTrue(plan.confirmation_required)
        self.assertIn("merge/split create new documents", " ".join(plan.risks))
        self.assertIn("合并 3 个 DOCX", plan.target)

        with self.assertRaisesRegex(GuardError, "explicit human confirmation required"):
            self.editor.execute(plan.id)
        self.assertEqual({p.name: digest(p) for p in (first, second, third)}, digests)
        self.assertFalse(any(self.root.glob("*.agent-*")))

        verified = self.editor.execute(plan.id, confirmed=True)
        output = Path(verified["output"])
        self.assertEqual(output.name, f"one.agent-{plan.id[:8]}.docx")
        merged = Document(output)
        texts = [paragraph.text for paragraph in merged.paragraphs if paragraph.text]
        self.assertEqual(texts, ["First", "alpha one", "Second", "beta two", "gamma three"])
        self.assertEqual(len(merged.sections), 3)          # each source kept its own section
        self.assertEqual(len(merged.tables), 1)
        self.assertEqual(merged.tables[0].cell(0, 0).text, "one.docx-a")
        self.assertEqual(verified["kind"], "docx_merge")
        self.assertEqual([entry["name"] for entry in verified["sources"]],
                         ["one.docx", "two.docx", "three.docx"])
        self.assertTrue(Path(verified["source_backups"]).is_dir())
        self.assertEqual(len(list(Path(verified["source_backups"]).iterdir())), 3)
        # Not one source byte changed, and the anchor's own backup is the pre-merge copy.
        self.assertEqual({p.name: digest(p) for p in (first, second, third)}, digests)
        self.assertEqual(digest(Path(verified["backup"])), digests["one.docx"])

        rolled_back = self.editor.rollback(verified["id"], confirmed=True)
        self.assertEqual(rolled_back["status"], "rolled_back")
        self.assertFalse(output.exists())
        self.assertEqual({p.name: digest(p) for p in (first, second, third)}, digests)

    def test_docx_merge_refuses_paths_that_are_not_file_names_in_the_directory(self):
        self._merge_source("plain.docx", ["plain"])
        outside = self.root.parent / "outside-merge.docx"
        document = Document(); document.add_paragraph("outside"); document.save(outside)
        self.addCleanup(outside.unlink, True)
        for sources in (["plain.docx", "..\\outside-merge.docx"],
                        ["plain.docx", "sub\\plain.docx"],
                        ["plain.docx", str(outside)],
                        ["plain.docx", "plain.docx"],
                        ["plain.docx", "other.txt"],
                        ["plain.docx"]):
            with self.subTest(sources=sources):
                with self.assertRaises(ValueError):
                    DocxMerge(path=str(self.root / "plain.docx"), sources=sources)
        self._merge_source("other.docx", ["other"])
        self._merge_source("second.docx", ["second"])
        with self.assertRaisesRegex(GuardError, "authorized file must be one of the merge sources"):
            self.editor.plan(DocxMerge(path=str(self.root / "plain.docx"),
                                       sources=["other.docx", "second.docx"]))
        # A name that does not exist is refused by the same path guard as any other tool.
        with self.assertRaisesRegex(GuardError, "could not be resolved safely"):
            self.editor.plan(DocxMerge(path=str(self.root / "plain.docx"),
                                       sources=["plain.docx", "missing.docx"]))
        self.assertFalse(self.editor.audit.exists())

    def test_docx_merge_refuses_when_a_source_changed_after_the_preview(self):
        first = self._merge_source("race-a.docx", ["alpha"])
        second = self._merge_source("race-b.docx", ["beta"])
        plan = self.editor.plan(DocxMerge(path=str(first), sources=["race-a.docx", "race-b.docx"]))
        document = Document(second)
        document.add_paragraph("changed after preview")
        document.save(second)
        with self.assertRaisesRegex(GuardError, "merge inputs changed since preview"):
            self.editor.execute(plan.id, confirmed=True)
        self.assertFalse(any(self.root.glob("*.agent-*")))
        self.assertTrue(Path(plan.id) and (self.editor.audit / plan.id).is_dir())

    def test_docx_split_writes_one_new_document_per_heading_and_rolls_back(self):
        source = self.root / "split.docx"
        document = Document()
        document.add_heading("Chapter one", level=1)
        document.add_paragraph("first body")
        document.add_heading("Chapter two", level=1)
        document.add_paragraph("second body")
        document.add_heading("Chapter three", level=1)
        document.add_paragraph("third body")
        document.save(source)
        before = source.read_bytes()
        plan = self.editor.plan(DocxSplit(path=str(source), level=1))
        self.assertTrue(plan.confirmation_required)
        self.assertIn("拆分为 3 个新 DOCX", plan.target)

        with self.assertRaisesRegex(GuardError, "explicit human confirmation required"):
            self.editor.execute(plan.id)
        self.assertEqual(source.read_bytes(), before)

        verified = self.editor.execute(plan.id, confirmed=True)
        outputs = [Path(path) for path in verified["outputs"]]
        self.assertEqual(len(outputs), 3)
        self.assertEqual([path.name for path in outputs],
                         [f"split.part-{index:02d}.agent-{plan.id[:8]}.docx" for index in (1, 2, 3)])
        parts = [[paragraph.text for paragraph in Document(path).paragraphs if paragraph.text]
                 for path in outputs]
        self.assertEqual(parts, [["Chapter one", "first body"], ["Chapter two", "second body"],
                                 ["Chapter three", "third body"]])
        self.assertEqual(verified["part_titles"], ["Chapter one", "Chapter two", "Chapter three"])
        self.assertEqual(source.read_bytes(), before)          # the source is untouched

        rolled_back = self.editor.rollback(verified["id"], confirmed=True)
        self.assertEqual(rolled_back["status"], "rolled_back")
        self.assertTrue(all(not path.exists() for path in outputs))
        self.assertEqual(source.read_bytes(), before)

    def test_docx_split_refuses_when_there_is_nothing_to_split(self):
        plain = self.root / "no-headings.docx"
        document = Document(); document.add_paragraph("just text"); document.save(plain)
        with self.assertRaisesRegex(GuardError, "只会得到 1 个部分"):
            self.editor.plan(DocxSplit(path=str(plain), level=1))
        deeper = self.root / "deep-headings.docx"
        document = Document(); document.add_heading("Sub", level=2); document.add_paragraph("x")
        document.save(deeper)
        with self.assertRaisesRegex(GuardError, "只会得到 1 个部分"):
            self.editor.plan(DocxSplit(path=str(deeper), level=1))
        self.assertFalse(self.editor.audit.exists())

    def test_merge_and_split_refuse_to_overwrite_an_existing_output(self):
        first = self._merge_source("taken-a.docx", ["alpha"])
        self._merge_source("taken-b.docx", ["beta"])
        plan_inputs = ["taken-a.docx", "taken-b.docx"]
        plan = self.editor.plan(DocxMerge(path=str(first), sources=plan_inputs))
        taken = self.root / f"taken-a.agent-{plan.id[:8]}.docx"
        taken.write_bytes(b"already here")
        digests = {path.name: digest(path) for path in (first, self.root / "taken-b.docx")}
        with self.assertRaisesRegex(GuardError, "output already exists"):
            self.editor.execute(plan.id, confirmed=True)
        self.assertEqual(taken.read_bytes(), b"already here")
        self.assertEqual({path.name: digest(path) for path in (first, self.root / "taken-b.docx")},
                         digests)

    def test_creation_rollback_refuses_a_tampered_audit_record(self):
        first = self._merge_source("tamper-a.docx", ["alpha"])
        self._merge_source("tamper-b.docx", ["beta"])
        plan = self.editor.plan(DocxMerge(path=str(first), sources=["tamper-a.docx", "tamper-b.docx"]))
        verified = self.editor.execute(plan.id, confirmed=True)
        output = Path(verified["output"])
        bystander = self.root / "bystander.docx"
        bystander.write_bytes(b"keep me")
        record = self.editor.audit / plan.id / "audit.json"
        log = json.loads(record.read_text(encoding="utf-8"))
        log["outputs"] = [str(bystander)]
        record.write_text(json.dumps(log, ensure_ascii=False, indent=2), encoding="utf-8")
        with self.assertRaisesRegex(GuardError, "audit output paths do not match original plan"):
            self.editor.rollback(plan.id, confirmed=True)
        self.assertEqual(bystander.read_bytes(), b"keep me")
        self.assertTrue(output.exists())

    def test_merge_and_split_carry_images_through_relationship_remapping(self):
        """Copied XML keeps r:embed ids, so the relationships must be re-created explicitly."""
        image = self.root / "pixel.png"
        image.write_bytes(_PIXEL_PNG)
        illustrated = self.root / "illustrated.docx"
        document = Document()
        document.add_heading("With picture", level=1)
        document.add_paragraph("before the picture")
        document.add_picture(str(image))
        document.add_heading("Without picture", level=1)
        document.add_paragraph("second part")
        document.save(illustrated)
        plain = self._merge_source("plain.docx", ["plain text"])

        merged = self.editor.plan(DocxMerge(path=str(illustrated),
                                            sources=["illustrated.docx", "plain.docx"]))
        verified = self.editor.execute(merged.id, confirmed=True)
        output = Path(verified["output"])
        reopened = Document(output)
        self.assertEqual(len(reopened.inline_shapes), 1)
        reloaded = Document(output)
        self.assertEqual(len(reloaded.part.rels), len([rel for rel in reloaded.part.rels.values()]))
        with zipfile.ZipFile(output) as archive:
            media = [name for name in archive.namelist() if name.startswith("word/media/")]
        self.assertEqual(len(media), 1)                     # the image part travelled with it
        self.assertGreaterEqual(verified["copy_report"]["relationships_remapped"], 1)
        self.editor.rollback(verified["id"], confirmed=True)

        split = self.editor.plan(DocxSplit(path=str(illustrated), level=1))
        split_verified = self.editor.execute(split.id, confirmed=True)
        parts = [Document(path) for path in split_verified["outputs"]]
        self.assertEqual(sum(len(part.inline_shapes) for part in parts), 1)
        self.editor.rollback(split_verified["id"], confirmed=True)

    def test_word_writes_do_not_invent_header_or_footer_parts(self):
        """Reading section.header creates an empty part; a write must not add one."""
        source = self.root / "no-header.docx"
        document = Document()
        document.add_paragraph("price here")
        document.save(source)
        with zipfile.ZipFile(source) as archive:
            self.assertEqual([name for name in archive.namelist()
                              if "header" in name or "footer" in name], [])
        plan = self.editor.plan(WordReplace(path=str(source), old="price", new="amount",
                                            paragraph_index=0))
        output = Path(self.editor.execute(plan.id, confirmed=True)["output"])
        with zipfile.ZipFile(output) as archive:
            self.assertEqual([name for name in archive.namelist()
                              if "header" in name or "footer" in name], [])
        self.assertEqual(Document(output).paragraphs[0].text, "amount here")

    def test_word_writes_keep_an_existing_header_and_its_text(self):
        source = self.root / "with-header.docx"
        document = Document()
        document.add_paragraph("price here")
        document.sections[0].header.paragraphs[0].text = "brand"
        document.save(source)
        plan = self.editor.plan(WordReplace(path=str(source), old="price", new="amount",
                                            paragraph_index=0))
        output = Path(self.editor.execute(plan.id, confirmed=True)["output"])
        with zipfile.ZipFile(output) as archive:
            self.assertIn("word/header1.xml", archive.namelist())
        self.assertEqual(Document(output).sections[0].header.paragraphs[0].text, "brand")
        self.assertEqual(Document(output).paragraphs[0].text, "amount here")

    def test_word_spanning_runs_preserves_style(self):
        source = self.root / "a.docx"
        doc = Document()
        p = doc.add_paragraph()
        p.add_run("Before Hel").bold = True
        p.add_run("lo world").italic = True
        doc.add_paragraph("untouched")
        doc.save(source)
        orig = source.read_bytes()
        plan = self.editor.plan(WordReplace(kind="word_replace", path=str(source), old="Hello", new="你好", paragraph_index=0))
        self.assertEqual(plan.before, "Before Hello world")
        record = self.editor.execute(plan.id)
        edited = Document(record["output"])
        self.assertEqual(edited.paragraphs[0].text, "Before 你好 world")
        self.assertTrue(edited.paragraphs[0].runs[0].bold)
        self.assertTrue(edited.paragraphs[0].runs[1].italic)
        self.assertEqual(edited.paragraphs[1].text, "untouched")
        self.assertEqual(source.read_bytes(), orig)
        self.assertEqual(Path(record["backup"]).read_bytes(), orig)

    def test_word_format_one_run_preserves_other_runs_and_undo(self):
        source = self.root / "format.docx"
        doc = Document()
        paragraph = doc.add_paragraph()
        paragraph.add_run("left").bold = True
        target = paragraph.add_run("target")
        target.italic = True
        paragraph.add_run("right").underline = True
        doc.add_paragraph("other paragraph")
        table = doc.add_table(rows=1, cols=1)
        table.cell(0, 0).text = "other cell"
        doc.save(source)
        original = source.read_bytes()
        plan = self.editor.plan(WordFormat(kind="word_format", path=str(source),
                                           paragraph_index=0, run_index=1,
                                           bold=True, italic=False, underline=True))
        self.assertEqual(plan.target, "paragraph[0].run[1]")
        self.assertIn("'italic': True", plan.before)
        self.assertIn("'underline': True", plan.after)
        record = self.editor.execute(plan.id)
        edited = Document(record["output"])
        runs = edited.paragraphs[0].runs
        self.assertEqual([(r.text, r.bold, r.italic, r.underline) for r in runs],
                         [("left", True, None, None), ("target", True, False, True),
                          ("right", None, None, True)])
        self.assertEqual(edited.paragraphs[1].text, "other paragraph")
        self.assertEqual(edited.tables[0].cell(0, 0).text, "other cell")
        self.assertEqual(Path(record["backup"]).read_bytes(), original)
        self.assertEqual(source.read_bytes(), original)
        self.assertEqual(json.loads((self.editor.audit / plan.id / "audit.json").read_text(encoding="utf8"))["diff"]["target"], plan.target)
        self.assertEqual(self.editor.rollback(plan.id, confirmed=True)["status"], "rolled_back")
        self.assertFalse(Path(record["output"]).exists())

    def test_word_table_cell_replacement_preserves_runs_other_cells(self):
        source = self.root / "table.docx"
        doc = Document()
        doc.add_paragraph("outside")
        table = doc.add_table(rows=2, cols=2)
        paragraph = table.cell(0, 0).paragraphs[0]
        paragraph.add_run("prefix Hel").bold = True
        paragraph.add_run("lo suffix").italic = True
        table.cell(0, 1).text = "unchanged neighbor"
        table.cell(1, 0).text = "unchanged next row"
        table.cell(1, 1).text = "unchanged bottom"
        doc.save(source)
        original = source.read_bytes()
        plan = self.editor.plan(WordTableCellReplace(kind="word_table_cell_replace",
            path=str(source), table_index=0, row_index=0, cell_index=0,
            paragraph_index=0, old="Hello", new="world"))
        self.assertEqual(plan.before, "prefix Hello suffix")
        self.assertEqual(plan.after, "prefix world suffix")
        self.assertEqual(plan.target, "table[0].row[0].cell[0].paragraph[0]")
        record = self.editor.execute(plan.id)
        edited = Document(record["output"])
        runs = edited.tables[0].cell(0, 0).paragraphs[0].runs
        self.assertEqual([(r.text, r.bold, r.italic) for r in runs],
                         [("prefix world", True, None), (" suffix", None, True)])
        self.assertEqual(edited.tables[0].cell(0, 1).text, "unchanged neighbor")
        self.assertEqual(edited.tables[0].cell(1, 0).text, "unchanged next row")
        self.assertEqual(edited.tables[0].cell(1, 1).text, "unchanged bottom")
        self.assertEqual(edited.paragraphs[0].text, "outside")
        self.assertEqual(source.read_bytes(), original)
        self.assertEqual(Path(record["backup"]).read_bytes(), original)
        self.assertEqual(record["status"], "verified")

    def test_word_new_edits_reject_invalid_types_and_ambiguous_targets(self):
        source = self.root / "invalid.docx"
        doc = Document()
        doc.add_paragraph("repeat repeat")
        table = doc.add_table(rows=1, cols=1)
        table.cell(0, 0).text = "repeat repeat"
        doc.save(source)
        with self.assertRaises(ValidationError):
            WordFormat(kind="word_format", path=str(source), paragraph_index=0,
                       run_index=0, bold="yes")
        with self.assertRaises(ValidationError):
            WordFormat(kind="word_format", path=str(source), paragraph_index=0, run_index=0)
        with self.assertRaises(ValidationError):
            WordTableCellReplace(kind="word_table_cell_replace", path=str(source),
                table_index=0, row_index=True, cell_index=0, paragraph_index=0,
                old="repeat", new="new")
        with self.assertRaisesRegex(GuardError, "exactly one occurrence"):
            self.editor.plan(WordTableCellReplace(kind="word_table_cell_replace",
                path=str(source), table_index=0, row_index=0, cell_index=0,
                paragraph_index=0, old="repeat", new="new"))
        with self.assertRaisesRegex(GuardError, "run index out of bounds"):
            self.editor.plan(WordFormat(kind="word_format", path=str(source),
                paragraph_index=0, run_index=3, bold=True))

    def test_existing_replace_checks_only_touched_run_xml(self):
        source = self.root / "run-xml.docx"
        doc = Document()
        paragraph = doc.add_paragraph()
        paragraph.add_run("replace")
        paragraph.add_run("later").add_tab()
        doc.save(source)
        plan = self.editor.plan(WordReplace(kind="word_replace", path=str(source),
                                             paragraph_index=0, old="replace", new="changed"))
        result = self.editor.execute(plan.id)
        self.assertEqual(Document(result["output"]).paragraphs[0].text, "changedlater\t")
        with self.assertRaisesRegex(GuardError, "non-text XML"):
            self.editor.plan(WordReplace(kind="word_replace", path=str(source),
                paragraph_index=0, old="later", new="new"))

    def test_word_new_edits_reject_hyperlinks_and_non_text_runs(self):
        source = self.root / "complex.docx"
        doc = Document()
        paragraph = doc.add_paragraph()
        paragraph.add_run("safe")
        link = OxmlElement("w:hyperlink")
        link.set(qn("w:anchor"), "bookmark")
        link_run = OxmlElement("w:r")
        text = OxmlElement("w:t")
        text.text = "linked"
        link_run.append(text)
        link.append(link_run)
        paragraph._p.append(link)
        table = doc.add_table(rows=1, cols=1)
        cell_paragraph = table.cell(0, 0).paragraphs[0]
        cell_paragraph.add_run("before")
        cell_paragraph.add_run("after").add_tab()
        doc.save(source)
        with self.assertRaisesRegex(GuardError, "hyperlink/unsupported"):
            self.editor.plan(WordFormat(kind="word_format", path=str(source),
                                        paragraph_index=0, run_index=0, bold=True))
        with self.assertRaisesRegex(GuardError, "non-text run XML"):
            self.editor.plan(WordTableCellReplace(kind="word_table_cell_replace",
                path=str(source), table_index=0, row_index=0, cell_index=0,
                paragraph_index=0, old="before", new="changed"))

    def test_excel_formula_and_style_preserved(self):
        source = self.root / "a.xlsx"
        wb = Workbook(); ws = wb.active; ws.title = "Sheet"
        ws["A1"] = 3; ws["B1"] = "=A1*2"; ws["B1"].font = Font(bold=True)
        wb.save(source)
        plan = self.editor.plan(ExcelSet(kind="excel_set", path=str(source), sheet="Sheet", cell="A1", value=5))
        record = self.editor.execute(plan.id)
        new = load_workbook(record["output"], data_only=False)
        self.assertEqual(new.active["B1"].value, "=A1*2")
        self.assertTrue(new.active["B1"].font.bold)
        self.assertEqual(new.active["A1"].value, 5)
        self.assertIsNone(load_workbook(record["output"], data_only=True).active["B1"].value)

    def test_formula_fill_preview_confirmation_preserves_style_backup_and_rollback(self):
        source = self.root / "fill.xlsx"
        wb = Workbook(); ws = wb.active; ws.title = "Data"
        for row in range(1, 4):
            ws[f"A{row}"] = row * 3
        ws["B1"] = "=(A1+2)*2"; ws["B2"].font = Font(bold=True)
        ws["C2"] = "untouched"
        wb.save(source)
        original = source.read_bytes()
        request = ExcelFormulaFill(kind="excel_formula_fill", path=str(source), sheet="Data",
                                   source_cell="B1", target_range="B2:B3", expected_source_formula="=(A1+2)*2")
        plan = self.editor.plan(request)
        self.assertEqual(plan.target, "Data!B2:B3")
        self.assertEqual(plan.after, repr({"B2": "=(A2+2)*2", "B3": "=(A3+2)*2"}))
        self.assertTrue(plan.confirmation_required)
        with self.assertRaisesRegex(GuardError, "human confirmation"):
            self.editor.execute(plan.id)
        self.assertFalse(self.editor.audit.exists())
        record = self.editor.execute(plan.id, confirmed=True)
        result = load_workbook(record["output"], data_only=False)
        self.assertEqual([result.active[f"B{row}"].value for row in range(1, 4)],
                         ["=(A1+2)*2", "=(A2+2)*2", "=(A3+2)*2"])
        self.assertTrue(result.active["B2"].font.bold)
        self.assertEqual(result.active["C2"].value, "untouched")
        result.close()
        cached = load_workbook(record["output"], data_only=True)
        self.assertIsNone(cached.active["B2"].value)
        cached.close()
        self.assertEqual(source.read_bytes(), original)
        self.assertEqual(Path(record["backup"]).read_bytes(), original)
        self.assertEqual(json.loads((self.editor.audit / plan.id / "audit.json").read_text(encoding="utf8"))["diff"]["after"], plan.after)
        self.assertEqual(self.editor.rollback(plan.id, confirmed=True)["status"], "rolled_back")
        self.assertFalse(Path(record["output"]).exists())

    def test_formula_fill_rejects_unsafe_formulas_ranges_and_occupied_targets(self):
        source = self.root / "fill-invalid.xlsx"
        wb = Workbook(); ws = wb.active; ws.title = "Data"
        ws["B1"] = "=A1*2"; ws["B2"] = "occupied"
        wb.save(source)
        base = dict(kind="excel_formula_fill", path=str(source), sheet="Data",
                    source_cell="B1", target_range="B2:B3", expected_source_formula="=A1*2")
        # Spaces around operators are safe; whitespace between references is not.
        ExcelFormulaFill(**{**base, "expected_source_formula": "=A1 + 2"})
        for change in (dict(target_range="B3:B4"), dict(target_range="B2:C3"),
                       dict(target_range="B2:B22"), dict(target_range="A2:A3"),
                       dict(expected_source_formula="=HYPERLINK(A1)"),
                       # ITERATION_PLANS round 2 deliberately made allowlisted functions
                       # (SUM/IF/ROUND/...) valid fill sources, so =SUM(A1:A2) is no longer
                       # in this list; the function-specific rejections live in
                       # test_formula_fill_function_allowlist_and_rejections.
                       dict(expected_source_formula="=A1+"),
                       dict(expected_source_formula="=A1**2"),
                       dict(expected_source_formula="=A1 A2"),
                       dict(expected_source_formula="=A1:A2"),
                       dict(expected_source_formula="=)"),
                       dict(expected_source_formula="=A1+)"),
                       dict(expected_source_formula="='Other'!A1")):
            with self.subTest(change=change), self.assertRaises(ValidationError):
                ExcelFormulaFill(**{**base, **change})
        with self.assertRaisesRegex(GuardError, "empty, unmerged"):
            self.editor.plan(ExcelFormulaFill(**base))
        ws["B2"] = None; wb.save(source)
        with self.assertRaisesRegex(GuardError, "expected_source_formula"):
            self.editor.plan(ExcelFormulaFill(**{**base, "expected_source_formula": "=A1+2"}))
        ws["B1"] = "=HYPERLINK(A1)"; wb.save(source)
        with self.assertRaisesRegex(ValidationError, "function|only local single-cell references"):
            ExcelFormulaFill(**{**base, "expected_source_formula": "=HYPERLINK(A1)"})
        ws["B1"] = "=A1*2"; wb.create_sheet("Extra"); wb.save(source)
        with self.assertRaisesRegex(GuardError, "one sheet"):
            self.editor.plan(ExcelFormulaFill(**base))
        self.assertFalse(self.editor.audit.exists())

    def test_formula_fill_validation_errors_state_the_expected_shape(self):
        """A local 7B model could not infer the range format from a bare rejection, so
        each message must name the expected shape and the value that was rejected."""
        base = {"path": "book.xlsx", "sheet": "Data", "source_cell": "B1",
                "expected_source_formula": "=A1*3"}
        # A bare single-cell target is unambiguous and normalised, not rejected.
        self.assertEqual(ExcelFormulaFill(**base, target_range="B2").target_range, "B2:B2")
        with self.assertRaisesRegex(ValueError, r"explicit bounded range like 'B2:B2'"):
            ExcelFormulaFill(**base, target_range="B2 or B3")
        with self.assertRaisesRegex(ValueError, r"immediately below B1"):
            ExcelFormulaFill(**base, target_range="C2:C2")
        with self.assertRaisesRegex(ValueError, r"invalid cell address 'Data!B1'"):
            ExcelFormulaFill(**{**base, "source_cell": "Data!B1"}, target_range="B2:B2")

    def test_normalising_a_one_cell_target_keeps_every_semantic_check(self):
        base = {"path": "book.xlsx", "sheet": "Data", "expected_source_formula": "=A1*3"}
        # Normalisation must not let a single cell escape the adjacency rules.
        with self.assertRaisesRegex(ValueError, r"immediately below B1"):
            ExcelFormulaFill(**base, source_cell="B1", target_range="B3")
        with self.assertRaisesRegex(ValueError, r"immediately below B1"):
            ExcelFormulaFill(**base, source_cell="B1", target_range="C2")
        with self.assertRaisesRegex(ValueError, r"immediately below B1"):
            ExcelFormulaFill(**base, source_cell="B1", target_range="A2")

    def test_each_write_request_has_exactly_one_legal_kind_so_it_defaults(self):
        """Requiring the model to echo a single-valued discriminator only added a way to
        fail (observed sending kind='XLSX' and kind='word_replace' at the batch tool)."""
        request = ExcelFormulaFill(path="book.xlsx", sheet="Data", source_cell="B1",
                                   target_range="B2:B2", expected_source_formula="=A1*3")
        self.assertEqual(request.kind, "excel_formula_fill")
        self.assertEqual(WordReplace(path="d.docx", paragraph_index=0, old="a", new="b").kind,
                         "word_replace")
        # A stray label is inert data, not a second failure mode: the dispatcher writes the
        # canonical kind into the plan, and the tool name is what selects the rules.
        self.assertEqual(ExcelFormulaFill(**{**request.model_dump(), "kind": "XLSX"}).kind, "XLSX")
        with self.assertRaises(ValidationError):
            ExcelFormulaFill(**{**request.model_dump(), "target_range": "not-a-range"})

    def test_formula_fill_preserves_absolute_references_and_refuses_overflow(self):
        source = self.root / "fill-absolute.xlsx"
        wb = Workbook(); ws = wb.active
        ws["B1"] = "=A$1+$A1+$A$1"
        wb.save(source)
        request = ExcelFormulaFill(kind="excel_formula_fill", path=str(source), sheet="Sheet",
            source_cell="B1", target_range="B2:B2", expected_source_formula="=A$1+$A1+$A$1")
        plan = self.editor.plan(request)
        self.assertEqual(plan.after, repr({"B2": "=A$1+$A2+$A$1"}))
        # Translating a relative reference past the last Excel row must fail closed.
        ws["B1048575"] = "=A1048576"
        wb.save(source)
        overflow = ExcelFormulaFill(kind="excel_formula_fill", path=str(source), sheet="Sheet",
            source_cell="B1048575", target_range="B1048576:B1048576",
            expected_source_formula="=A1048576")
        with self.assertRaisesRegex(GuardError, "unsafe or out-of-range"):
            self.editor.plan(overflow)
        self.assertFalse(self.editor.audit.exists())

    def test_formula_fill_function_allowlist_and_rejections(self):
        """Round 2: allowlisted deterministic functions are valid fill sources; volatile,
        external, unknown functions and cross-sheet references stay rejected by name."""
        source = self.root / "fill-functions.xlsx"
        wb = Workbook(); ws = wb.active; ws.title = "Data"
        ws["A1"], ws["A2"], ws["A3"] = 1, 2, 3
        wb.save(source)

        accepted = ("=SUM(A1:A3)", "=AVERAGE(A1:A3)", "=MAX(A1:A3)*2", "=COUNT(A1:A3)",
                    "=IF(A1>0,\"yes\",\"no\")", "=ROUND(A1/3,2)", "=MIN(A1,A2)",
                    "=IF(AND(A1>0,A2<5),ROUND(A1,2),0)")
        for formula in accepted:
            with self.subTest(formula=formula):
                request = ExcelFormulaFill(path=str(source), sheet="Data", source_cell="B1",
                                           target_range="B2:B2", expected_source_formula=formula)
                self.assertEqual(request.expected_source_formula, formula)

        rejected = {
            "=NOW()": "volatile",
            "=TODAY()": "volatile",
            "=RAND()": "volatile",
            "=OFFSET(A1,1,1)": "volatile",
            "=INDIRECT(\"A1\")": "volatile",
            "=ROW()": "volatile",
            "=HYPERLINK(A1)": "outside the workbook",
            "=VLOOKUP(A1,A1:B3,2)": "not in the allowed list",
            "=SUM(Sheet2!A1:A3)": "cross-sheet",
            "='Other'!A1": "cross-sheet",
            "=SUM(A1:A3": "incomplete",
            "=SUM(,A1)": "comma must separate",
            "=SUM(A1:A1048577)": "exceeds Excel limits",
            "=SUM(A1:A3) A2": "only local references",
        }
        for formula, reason in rejected.items():
            with self.subTest(formula=formula):
                with self.assertRaisesRegex(ValidationError, reason):
                    ExcelFormulaFill(path=str(source), sheet="Data", source_cell="B1",
                                     target_range="B2:B2", expected_source_formula=formula)

    def test_formula_fill_function_round_trip_verifies_and_rolls_back(self):
        source = self.root / "fill-sum.xlsx"
        wb = Workbook(); ws = wb.active; ws.title = "Data"
        ws["A1"], ws["A2"], ws["A3"] = 1, 2, 3
        ws["B1"] = "=SUM(A1:A3)"
        wb.save(source)
        original = source.read_bytes()

        plan = self.editor.plan(ExcelFormulaFill(
            path=str(source), sheet="Data", source_cell="B1", target_range="B2:B2",
            expected_source_formula="=SUM(A1:A3)"))
        self.assertTrue(plan.confirmation_required)
        self.assertIn("not calculated", " ".join(plan.risks))
        with self.assertRaisesRegex(GuardError, "human confirmation"):
            self.editor.execute(plan.id)
        record = self.editor.execute(plan.id, confirmed=True)
        reopened = load_workbook(record["output"], data_only=False)
        self.assertEqual(reopened["Data"]["B2"].value, "=SUM(A2:A4)")
        reopened.close()
        # The formula text is written; no cached result is produced or claimed as calculated.
        cached = load_workbook(record["output"], data_only=True)
        self.assertIsNone(cached["Data"]["B2"].value)
        cached.close()
        self.assertEqual(source.read_bytes(), original)
        self.assertEqual(self.editor.rollback(record["id"], confirmed=True)["status"], "rolled_back")
        self.assertFalse(Path(record["output"]).exists())

    def test_formula_fill_to_the_right_and_direction_validation(self):
        source = self.root / "fill-right.xlsx"
        wb = Workbook(); ws = wb.active; ws.title = "Data"
        ws["A1"], ws["A2"] = 10, 20
        ws["B1"] = "=A1*2"
        wb.save(source)
        original = source.read_bytes()

        plan = self.editor.plan(ExcelFormulaFill(
            path=str(source), sheet="Data", source_cell="B1", target_range="C1:D1",
            expected_source_formula="=A1*2", direction="right"))
        self.assertTrue(plan.confirmation_required)
        record = self.editor.execute(plan.id, confirmed=True)
        reopened = load_workbook(record["output"], data_only=False)
        self.assertEqual(reopened["Data"]["C1"].value, "=B1*2")
        self.assertEqual(reopened["Data"]["D1"].value, "=C1*2")
        self.assertEqual(reopened["Data"]["B1"].value, "=A1*2")  # source untouched
        reopened.close()
        self.assertEqual(source.read_bytes(), original)

        # Wrong direction, or a target that is not adjacent, must be refused.
        for kwargs, message in (
                (dict(target_range="C1:D1", direction="down"), "immediately below"),
                (dict(target_range="B2:B3", direction="right"), "immediately to the right"),
                (dict(target_range="C2:D2", direction="right"), "immediately to the right"),
                (dict(target_range="C1:C21", direction="right"), "immediately to the right")):
            with self.subTest(**kwargs):
                with self.assertRaisesRegex(ValueError, message):
                    ExcelFormulaFill(path=str(source), sheet="Data", source_cell="B1",
                                     expected_source_formula="=A1*2", **kwargs)

    def test_formula_fill_right_refuses_an_occupied_target(self):
        source = self.root / "fill-right-occupied.xlsx"
        wb = Workbook(); ws = wb.active; ws.title = "Data"
        ws["A1"] = 10
        ws["B1"] = "=A1*2"
        ws["C1"] = "already here"
        wb.save(source)
        original = source.read_bytes()
        request = ExcelFormulaFill(path=str(source), sheet="Data", source_cell="B1",
                                   target_range="C1:D1", expected_source_formula="=A1*2",
                                   direction="right")
        with self.assertRaisesRegex(GuardError, "never overwriting existing content"):
            self.editor.plan(request)
        self.assertEqual(source.read_bytes(), original)
        self.assertFalse(self.editor.audit.exists())

    def test_formula_fill_rejects_source_changed_after_preview(self):
        source = self.root / "fill-changed.xlsx"
        wb = Workbook(); ws = wb.active; ws["A1"] = 1; ws["B1"] = "=A1*2"; wb.save(source)
        request = ExcelFormulaFill(kind="excel_formula_fill", path=str(source), sheet="Sheet",
            source_cell="B1", target_range="B2:B2", expected_source_formula="=A1*2")
        plan = self.editor.plan(request)
        ws["A1"] = 3; wb.save(source)
        altered = source.read_bytes()
        with self.assertRaisesRegex(GuardError, "source changed since preview"):
            self.editor.execute(plan.id, confirmed=True)
        self.assertEqual(source.read_bytes(), altered)
        self.assertFalse(self.editor.audit.exists())

    def test_formula_fill_rejects_cached_results_and_unknown_parts(self):
        source = self.root / "fill-package.xlsx"
        wb = Workbook(); ws = wb.active; ws["B1"] = "=A1+1"; wb.save(source)
        request = ExcelFormulaFill(kind="excel_formula_fill", path=str(source), sheet="Sheet",
            source_cell="B1", target_range="B2:B2", expected_source_formula="=A1+1")
        with zipfile.ZipFile(source) as archive:
            parts = {name: archive.read(name) for name in archive.namelist()}
        sheet = parts["xl/worksheets/sheet1.xml"]
        self.assertIn(b"<v></v>", sheet)
        parts["xl/worksheets/sheet1.xml"] = sheet.replace(b"<v></v>", b"<v>2</v>")
        with zipfile.ZipFile(source, "w") as archive:
            for name, blob in parts.items():
                archive.writestr(name, blob)
        with self.assertRaisesRegex(GuardError, "cached results"):
            self.editor.plan(request)
        with zipfile.ZipFile(source, "a") as archive:
            archive.writestr("xl/unknown.xml", b"<unknown/>")
        with self.assertRaisesRegex(GuardError, "unknown parts"):
            self.editor.plan(request)
        self.assertFalse(self.editor.audit.exists())

    def test_formula_fill_failure_removes_temp_and_keeps_source(self):
        source = self.root / "fill-failure.xlsx"
        wb = Workbook(); wb.active["B1"] = "=A1+1"; wb.save(source)
        original = source.read_bytes()
        plan = self.editor.plan(ExcelFormulaFill(kind="excel_formula_fill", path=str(source),
            sheet="Sheet", source_cell="B1", target_range="B2:B2", expected_source_formula="=A1+1"))
        with patch("office_agent.writes.os.link", side_effect=OSError("forced failure")):
            with self.assertRaises(OSError):
                self.editor.execute(plan.id, confirmed=True)
        job = self.editor.audit / plan.id
        self.assertEqual(source.read_bytes(), original)
        self.assertEqual((job / "backup.xlsx").read_bytes(), original)
        self.assertFalse((job / "working.xlsx").exists())
        self.assertFalse(source.with_name(f"{source.stem}.agent-{plan.id[:8]}.xlsx").exists())
        self.assertEqual(json.loads((job / "audit.json").read_text(encoding="utf8"))["status"], "failed")

    def test_excel_trim_preserves_workbook_backup_audit_and_rollback(self):
        source = self.root / "trim.xlsx"
        wb = Workbook(); ws = wb.active; ws.title = "Data"
        ws["A1"] = "  hello \t"; ws["A1"].font = Font(bold=True)
        ws["B1"] = "untouched"; ws["C1"] = "=LEN(B1)"
        wb.save(source)
        original = source.read_bytes()
        plan = self.editor.plan(ExcelTrim(kind="excel_trim", path=str(source),
                                          sheet="Data", cell="A1", expected_original="  hello \t"))
        self.assertEqual((plan.target, plan.before, plan.after),
                         ("Data!A1", repr("  hello \t"), repr("hello")))
        record = self.editor.execute(plan.id)
        result = load_workbook(record["output"], data_only=False).active
        self.assertEqual((result["A1"].value, result["B1"].value, result["C1"].value),
                         ("hello", "untouched", "=LEN(B1)"))
        self.assertTrue(result["A1"].font.bold)
        self.assertEqual(source.read_bytes(), original)
        self.assertEqual(Path(record["backup"]).read_bytes(), original)
        self.assertEqual(json.loads((self.editor.audit / plan.id / "audit.json").read_text(encoding="utf8"))["diff"]["after"], repr("hello"))
        self.assertEqual(self.editor.rollback(plan.id, confirmed=True)["status"], "rolled_back")
        self.assertFalse(Path(record["output"]).exists())

    def test_excel_trim_rejects_ambiguous_formula_and_noop(self):
        source = self.root / "trim-invalid.xlsx"
        wb = Workbook(); ws = wb.active; ws.title = "Data"
        ws["A1"] = "  value  "; ws["B1"] = "=1+1"; ws["C1"] = "clean"
        ws["D1"] = "  =1+1  "
        ws["D1"].data_type = "s"
        wb.save(source)
        cases = (("A1", "wrong", "does not match the current value"),
                 ("B1", "=1+1", "formulas are never trimmed"),
                 ("C1", "clean", "would change nothing"),
                 ("D1", "  =1+1  ", "would produce a formula"))
        for cell, expected, specific in cases:
            with self.subTest(cell=cell):
                with self.assertRaises(GuardError) as caught:
                    self.editor.plan(ExcelTrim(kind="excel_trim", path=str(source),
                                               sheet="Data", cell=cell, expected_original=expected))
                message = str(caught.exception)
                self.assertRegex(message, specific)
                # Every rejection also states the remedy, so a model can correct itself.
                self.assertIn("read the target with read_excel_region first", message)
        with self.assertRaises(ValidationError):
            ExcelTrim(kind="excel_trim", path=str(source), sheet="Data", cell="A1:B1",
                      expected_original="  value  ")
        with self.assertRaises(ValidationError):
            ExcelTrim(kind="excel_trim", path=str(source), sheet="Data", cell="A1",
                      expected_original=4)
        with self.assertRaisesRegex(GuardError, "docx/xlsx"):
            doc = self.root / "wrong.docx"; Document().save(doc)
            self.editor.plan(ExcelTrim(kind="excel_trim", path=str(doc),
                                       sheet="Data", cell="A1", expected_original="  value  "))
        self.assertFalse(self.editor.audit.exists())

    def test_excel_trim_rejects_cached_formula_results(self):
        source = self.root / "trim-cached.xlsx"
        wb = Workbook(); ws = wb.active; ws.title = "Data"
        ws["A1"] = "  text "; ws["B1"] = "=1+1"; wb.save(source)
        with zipfile.ZipFile(source) as archive:
            parts = {name: archive.read(name) for name in archive.namelist()}
        sheet = parts["xl/worksheets/sheet1.xml"]
        self.assertIn(b"<v></v>", sheet)
        parts["xl/worksheets/sheet1.xml"] = sheet.replace(b"<v></v>", b"<v>2</v>")
        with zipfile.ZipFile(source, "w") as archive:
            for name, data in parts.items():
                archive.writestr(name, data)
        with self.assertRaisesRegex(GuardError, "cached results"):
            self.editor.plan(ExcelTrim(kind="excel_trim", path=str(source),
                                       sheet="Data", cell="A1", expected_original="  text "))

    def test_rejects_model_supplied_formula(self):
        from pydantic import ValidationError
        with self.assertRaisesRegex(ValidationError, "formula writes"):
            ExcelSet(kind="excel_set", path=str(self.root / "a.xlsx"), sheet="Sheet", cell="B2", value='=HYPERLINK("https://example.invalid")')

    def test_rejects_cached_formula_loss(self):
        source = self.root / "cached.xlsx"
        wb = Workbook(); wb.active["A1"] = 1; wb.active["B1"] = "=A1+1"; wb.save(source)
        # Simulate an existing calculated result: openpyxl drops cached values on save.
        with zipfile.ZipFile(source) as src:
            parts = {name: src.read(name) for name in src.namelist()}
        sheet = parts["xl/worksheets/sheet1.xml"]
        self.assertIn(b"<v></v>", sheet)
        parts["xl/worksheets/sheet1.xml"] = sheet.replace(b"<v></v>", b"<v>2</v>")
        with zipfile.ZipFile(source, "w") as out:
            for name, data in parts.items():
                out.writestr(name, data)
        with self.assertRaisesRegex(GuardError, "cached results"):
            self.editor.plan(ExcelSet(kind="excel_set", path=str(source), sheet="Sheet", cell="A1", value=2))

    def test_excel_bar_chart_confirm_reopen_backup_and_rollback(self):
        source = self.root / "chart.xlsx"
        wb = Workbook(); ws = wb.active; ws.title = "Data"
        ws.append(["A", 2, "unchanged"]); ws.append(["B", 3, "still here"])
        ws["C1"].font = Font(bold=True)
        wb.save(source)
        original = source.read_bytes()
        request = ExcelBarChart(kind="excel_bar_chart", path=str(source), sheet="Data",
                                values_range="B1:B2", categories_range="A1:A2",
                                anchor="A1", title="Sales")
        plan = self.editor.plan(request)
        self.assertTrue(plan.confirmation_required)
        self.assertEqual(plan.target, "Data!A1")
        with self.assertRaisesRegex(GuardError, "human confirmation"):
            self.editor.execute(plan.id)
        self.assertFalse(self.editor.audit.exists())
        record = self.editor.execute(plan.id, confirmed=True)
        result = load_workbook(record["output"])
        self.assertEqual([(result.active[f"A{i}"].value, result.active[f"B{i}"].value,
                           result.active[f"C{i}"].value) for i in (1, 2)],
                         [("A", 2, "unchanged"), ("B", 3, "still here")])
        self.assertTrue(result.active["C1"].font.bold)
        self.assertEqual(len(result.active._charts), 1)
        chart = result.active._charts[0]
        self.assertIsInstance(chart, BarChart)
        self.assertEqual(chart.series[0].val.numRef.f, "'Data'!$B$1:$B$2")
        self.assertEqual(chart.series[0].cat.numRef.f, "'Data'!$A$1:$A$2")
        self.assertEqual((chart.anchor._from.col, chart.anchor._from.row), (0, 0))
        result.close()
        self.assertEqual(source.read_bytes(), original)
        self.assertEqual(Path(record["backup"]).read_bytes(), original)
        self.assertEqual(json.loads((self.editor.audit / plan.id / "audit.json").read_text(encoding="utf8"))["status"], "verified")
        self.assertEqual(self.editor.rollback(plan.id, confirmed=True)["status"], "rolled_back")
        self.assertFalse(Path(record["output"]).exists())

    def test_excel_bar_chart_strict_ranges_and_source_guards(self):
        source = self.root / "chart-invalid.xlsx"
        wb = Workbook(); ws = wb.active; ws.title = "Data"
        ws["A1"] = "first"; ws["A2"] = "second"; ws["B1"] = 1; ws["B2"] = "=1+1"
        wb.save(source)
        base = dict(kind="excel_bar_chart", path=str(source), sheet="Data",
                    values_range="B1:B2", categories_range="A1:A2", title="Sales")
        for change in (dict(values_range="B1:C2"), dict(values_range="B1:B21"),
                       dict(categories_range="A1:A3"), dict(categories_range="B1:B2"),
                       dict(anchor="B2"), dict(title="x" * 101), dict(values_range="B1")):
            with self.subTest(change=change), self.assertRaises(ValidationError):
                ExcelBarChart(**{**base, **change})
        with self.assertRaisesRegex(GuardError, "numeric literal"):
            self.editor.plan(ExcelBarChart(**base))
        ws["B2"] = 4; ws["A2"] = "=1+1"; wb.save(source)
        with self.assertRaisesRegex(GuardError, "categories.*literal"):
            self.editor.plan(ExcelBarChart(**base))
        ws["A2"] = "second"; wb.create_sheet("Extra"); wb.save(source)
        with self.assertRaisesRegex(GuardError, "one sheet"):
            self.editor.plan(ExcelBarChart(**base))
        del wb["Extra"]
        chart = BarChart(); chart.add_data(Reference(ws, min_col=2, min_row=1, max_row=2))
        ws.add_chart(chart, "D1"); wb.save(source)
        with self.assertRaisesRegex(GuardError, "unsupported|charts/drawings"):
            self.editor.plan(ExcelBarChart(**base))
        self.assertFalse(self.editor.audit.exists())

    def test_excel_bar_chart_optional_categories_and_failed_verification(self):
        source = self.root / "chart-no-categories.xlsx"
        wb = Workbook(); ws = wb.active; ws.title = "Data"
        for row in range(1, 21):
            ws.cell(row, 1, row)
        wb.save(source)
        request = ExcelBarChart(kind="excel_bar_chart", path=str(source), sheet="Data",
                                values_range="A1:A20", title="Numbers")
        plan = self.editor.plan(request)
        with patch("office_agent.writes._chart_projection", side_effect=GuardError("tampered chart")):
            with self.assertRaisesRegex(GuardError, "tampered chart"):
                self.editor.execute(plan.id, confirmed=True)
        job = self.editor.audit / plan.id
        self.assertEqual(json.loads((job / "audit.json").read_text(encoding="utf8"))["status"], "failed")
        self.assertFalse((job / "working.xlsx").exists())
        self.assertFalse(source.with_name(f"{source.stem}.agent-{plan.id[:8]}.xlsx").exists())
        self.assertEqual((job / "backup.xlsx").read_bytes(), source.read_bytes())
        plan = self.editor.plan(request)
        record = self.editor.execute(plan.id, confirmed=True)
        result = load_workbook(record["output"])
        self.assertIsNone(result.active._charts[0].series[0].cat)
        self.assertEqual(result.active._charts[0].series[0].val.numRef.f, "'Data'!$A$1:$A$20")
        result.close()

    def _chart_book(self, name="chart-general.xlsx"):
        source = self.root / name
        wb = Workbook(); ws = wb.active; ws.title = "Data"
        ws.append(["Q1", 10, 5]); ws.append(["Q2", 20, 8]); ws.append(["Q3", 30, 11])
        ws["E1"] = "=1+1"
        wb.save(source)
        return source

    def _chart_and_reopen(self, request):
        plan = self.editor.plan(request)
        self.assertTrue(plan.confirmation_required)
        with self.assertRaisesRegex(GuardError, "human confirmation"):
            self.editor.execute(plan.id)
        record = self.editor.execute(plan.id, confirmed=True)
        workbook = load_workbook(record["output"])
        return workbook, workbook["Data"]._charts[0], record

    def test_excel_chart_multi_series_bar_keeps_every_series_and_labels(self):
        source = self._chart_book()
        original = source.read_bytes()
        workbook, chart, record = self._chart_and_reopen(ExcelChart(
            path=str(source), sheet="Data", chart_type="bar", series_ranges=["B1:B3", "C1:C3"],
            categories_range="A1:A3", title="Quarterly",
            category_axis_title="Quarter", value_axis_title="Units", legend_position="bottom"))
        self.assertIsInstance(chart, BarChart)
        self.assertEqual([s.val.numRef.f for s in chart.series],
                         ["'Data'!$B$1:$B$3", "'Data'!$C$1:$C$3"])
        self.assertEqual(chart.series[0].cat.numRef.f, "'Data'!$A$1:$A$3")
        self.assertEqual(chart.x_axis.title.tx.rich.p[0].r[0].t, "Quarter")
        self.assertEqual(chart.y_axis.title.tx.rich.p[0].r[0].t, "Units")
        self.assertEqual(chart.legend.position, "b")
        # Non-target data and the source file itself are untouched.
        self.assertEqual([workbook["Data"][f"A{i}"].value for i in (1, 2, 3)], ["Q1", "Q2", "Q3"])
        self.assertEqual(workbook["Data"]["E1"].value, "=1+1")
        workbook.close()
        self.assertEqual(source.read_bytes(), original)
        self.assertEqual(Path(record["backup"]).read_bytes(), original)
        self.assertEqual(self.editor.rollback(plan_id := record["id"], confirmed=True)["status"], "rolled_back")
        self.assertFalse(Path(record["output"]).exists())

    def test_excel_chart_line_and_pie_types(self):
        for chart_type, expected in (("line", LineChart), ("pie", PieChart)):
            with self.subTest(chart_type=chart_type):
                source = self._chart_book(f"chart-{chart_type}.xlsx")
                workbook, chart, record = self._chart_and_reopen(ExcelChart(
                    path=str(source), sheet="Data", chart_type=chart_type,
                    series_ranges=["B1:B3"], categories_range="A1:A3",
                    title=f"{chart_type} chart", legend_position="right"))
                self.assertIsInstance(chart, expected)
                self.assertEqual(len(chart.series), 1)
                self.assertEqual(chart.series[0].val.numRef.f, "'Data'!$B$1:$B$3")
                workbook.close()
                self.editor.rollback(record["id"], confirmed=True)

    def test_excel_chart_legend_can_be_removed(self):
        source = self._chart_book("chart-no-legend.xlsx")
        workbook, chart, record = self._chart_and_reopen(ExcelChart(
            path=str(source), sheet="Data", chart_type="bar", series_ranges=["B1:B3"],
            title="No legend", legend_position="none"))
        self.assertIsNone(chart.legend)
        workbook.close()
        self.editor.rollback(record["id"], confirmed=True)

    def test_excel_chart_rejects_a_formula_cell_as_a_data_source(self):
        source = self._chart_book("chart-formula.xlsx")
        original = source.read_bytes()
        for ranges in ({"series_ranges": ["E1:E3"]},
                       {"series_ranges": ["B1:B3", "E1:E3"]}):
            with self.subTest(ranges=ranges):
                with self.assertRaisesRegex(GuardError, "finite numeric literal cells"):
                    self.editor.plan(ExcelChart(path=str(source), sheet="Data",
                                                title="bad", **ranges))
        self.assertEqual(source.read_bytes(), original)
        self.assertFalse(self.editor.audit.exists())

    def test_excel_chart_rejects_inconsistent_or_axisless_requests(self):
        source = self._chart_book("chart-invalid-general.xlsx")
        cases = (
            ("pie chart takes exactly one data column",
             dict(chart_type="pie", series_ranges=["B1:B3", "C1:C3"], title="t")),
            ("pie chart has no axes",
             dict(chart_type="pie", series_ranges=["B1:B3"], title="t", value_axis_title="Units")),
            ("each series must be a different column",
             dict(series_ranges=["B1:B3", "B1:B3"], title="t")),
            ("every series must cover exactly the same rows",
             dict(series_ranges=["B1:B3", "C2:C3"], title="t")),
            ("category column must be distinct",
             dict(series_ranges=["B1:B3"], categories_range="B1:B3", title="t")),
            ("category column must be distinct",
             dict(series_ranges=["B1:B3", "C1:C3"], categories_range="C1:C3", title="t")),
        )
        for message, kwargs in cases:
            with self.subTest(message=message, kwargs=kwargs):
                with self.assertRaisesRegex(ValueError, message):
                    ExcelChart(path=str(source), sheet="Data", **kwargs)
        self.assertFalse(self.editor.audit.exists())

    def test_excel_chart_still_refuses_a_workbook_with_unknown_parts(self):
        """The general chart must inherit the same fail-closed package gate."""
        source = self._chart_book("chart-unknown-parts.xlsx")
        with zipfile.ZipFile(source) as archive:
            parts = {name: archive.read(name) for name in archive.namelist()}
        parts["xl/pivotCache/pivotCacheDefinition1.xml"] = b"<pivotCacheDefinition/>"
        with zipfile.ZipFile(source, "w", zipfile.ZIP_DEFLATED) as archive:
            for name, data in parts.items():
                archive.writestr(name, data)
        # Either gate may refuse it first; what matters is that it is refused, not written.
        with self.assertRaisesRegex(GuardError, "high-fidelity adapter|unknown parts"):
            self.editor.plan(ExcelChart(path=str(source), sheet="Data", title="blocked",
                                        series_ranges=["B1:B3"]))
        self.assertFalse(self.editor.audit.exists())

    def test_excel_bar_chart_rejects_nonempty_defined_names(self):
        source = self.root / "chart-defined-names.xlsx"
        wb = Workbook(); ws = wb.active; ws.title = "Data"
        ws["A1"] = 1; ws["A2"] = 2
        wb.defined_names.add(DefinedName("Revenue", attr_text="'Data'!$A$1:$A$2"))
        wb.save(source)
        request = ExcelBarChart(kind="excel_bar_chart", path=str(source), sheet="Data",
                                values_range="A1:A2", title="Numbers")
        with self.assertRaisesRegex(GuardError, "nonempty.*definedNames"):
            self.editor.plan(request)

    def test_excel_bar_chart_rejects_cached_formula_and_unknown_ooxml(self):
        source = self.root / "chart-cached.xlsx"
        wb = Workbook(); ws = wb.active; ws.title = "Data"
        ws["A1"] = 1; ws["A2"] = 2; ws["C1"] = "=1+1"; wb.save(source)
        request = ExcelBarChart(kind="excel_bar_chart", path=str(source), sheet="Data",
                                values_range="A1:A2", title="Safe")
        with zipfile.ZipFile(source) as archive:
            parts = {name: archive.read(name) for name in archive.namelist()}
        name = "xl/worksheets/sheet1.xml"
        self.assertIn(b"<v></v>", parts[name])
        parts[name] = parts[name].replace(b"<v></v>", b"<v>2</v>")
        with zipfile.ZipFile(source, "w") as archive:
            for name, blob in parts.items():
                archive.writestr(name, blob)
        with self.assertRaisesRegex(GuardError, "cached results"):
            self.editor.plan(request)
        with zipfile.ZipFile(source, "a") as archive:
            archive.writestr("xl/unknown.xml", b"<unknown/>")
        with self.assertRaisesRegex(GuardError, "unknown|unsupported|risky"):
            self.editor.plan(request)

    def test_batch_needs_confirmation(self):
        source = self.root / "a.docx"; doc = Document(); doc.add_paragraph("x x"); doc.save(source)
        plan = self.editor.plan(WordReplace(kind="word_replace", path=str(source), old="x", new="y", paragraph_index=0, expected_count=2))
        with self.assertRaises(GuardError):
            self.editor.execute(plan.id)
        self.assertEqual(Document(source).paragraphs[0].text, "x x")
        result = self.editor.execute(plan.id, confirmed=True)
        self.assertEqual(Document(result["output"]).paragraphs[0].text, "y y")

    def test_existing_output_is_not_deleted(self):
        source = self.root / "a.docx"; doc = Document(); doc.add_paragraph("old"); doc.save(source)
        plan = self.editor.plan(WordReplace(kind="word_replace", path=str(source), old="old", new="new", paragraph_index=0))
        collision = source.with_name(f"a.agent-{plan.id[:8]}.docx")
        collision.write_bytes(b"unrelated user file")
        original = source.read_bytes()
        with self.assertRaises(GuardError):
            self.editor.execute(plan.id)
        self.assertEqual(collision.read_bytes(), b"unrelated user file")
        self.assertEqual(source.read_bytes(), original)
        self.assertFalse((self.editor.audit / plan.id).exists(),
                         "output collision must not leave an unaudited empty job")

    def test_human_rollback_only_deletes_unchanged_verified_copy(self):
        source = self.root / "a.docx"; doc = Document(); doc.add_paragraph("old"); doc.save(source)
        plan = self.editor.plan(WordReplace(kind="word_replace", path=str(source), old="old", new="new", paragraph_index=0))
        record = self.editor.execute(plan.id)
        output = Path(record["output"])
        verified_bytes = output.read_bytes()
        with self.assertRaisesRegex(GuardError, "human confirmation"):
            self.editor.rollback(plan.id)
        self.assertTrue(output.exists())
        output.write_bytes(b"replaced by another process")
        with self.assertRaisesRegex(GuardError, "changed since publication"):
            self.editor.rollback(plan.id, confirmed=True)
        self.assertEqual(output.read_bytes(), b"replaced by another process")
        output.write_bytes(verified_bytes)
        undone = self.editor.rollback(plan.id, confirmed=True)
        self.assertEqual(undone["status"], "rolled_back")
        self.assertFalse(output.exists())
        self.assertEqual(Document(source).paragraphs[0].text, "old")
        self.assertTrue(Path(record["backup"]).exists())

    def test_failure_keeps_source_and_removes_temp(self):
        source = self.root / "a.docx"; doc = Document(); doc.add_paragraph("old"); doc.save(source)
        plan = self.editor.plan(WordReplace(kind="word_replace", path=str(source), old="old", new="new", paragraph_index=0))
        original = source.read_bytes()
        with patch("office_agent.writes.os.link", side_effect=OSError("forced failure")):
            with self.assertRaises(OSError):
                self.editor.execute(plan.id)
        job = self.editor.audit / plan.id
        self.assertEqual(source.read_bytes(), original)
        self.assertFalse((job / "working.docx").exists())
        self.assertEqual(json.loads((job / "audit.json").read_text(encoding="utf8"))["status"], "failed")
        self.assertEqual((job / "backup.docx").read_bytes(), original)


if __name__ == "__main__":
    unittest.main()
