import tempfile
import unittest
import zipfile
from pathlib import Path

from docx import Document
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference

from .documents import (
    OCRResult,
    classify_document,
    inspect_docx,
    inspect_file,
    inspect_xlsx,
    search_docx,
)


class DocumentInspectionTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.root = Path(self.tmp.name)

    def tearDown(self):
        self.tmp.cleanup()

    def test_docx_text_structure_and_chunked_search(self):
        path = self.root / "sample.docx"
        doc = Document()
        doc.add_heading("Quarterly report", level=1)
        doc.add_paragraph("alpha " * 20 + "needle")
        table = doc.add_table(rows=1, cols=2)
        table.cell(0, 0).text = "A"
        table.cell(0, 1).text = "B"
        doc.sections[0].header.paragraphs[0].text = "Header"
        doc.sections[0].footer.paragraphs[0].text = "Footer"
        doc.save(path)

        report = inspect_docx(path)
        self.assertTrue(report["readable"])
        self.assertTrue(report["paragraphs"][0]["is_heading"])
        self.assertEqual(report["tables"], [[['A', 'B']]])
        self.assertEqual(report["headers_footers"][0]["header"], ["Header"])
        self.assertTrue(search_docx(path, "needle", chunk_size=30, overlap=5))

    def test_xlsx_formula_cached_value_and_features(self):
        path = self.root / "sample.xlsx"
        wb = Workbook()
        ws = wb.active
        ws.title = "Data"
        ws["A1"] = 2
        ws["A2"] = 3
        ws["A3"] = "=SUM(A1:A2)"
        ws.merge_cells("B1:C1")
        chart = BarChart()
        chart.add_data(Reference(ws, min_col=1, min_row=1, max_row=2))
        ws.add_chart(chart, "E2")
        wb.save(path)

        report = inspect_xlsx(path)
        sheet = report["sheets"][0]
        formula = next(cell for cell in sheet["cells"] if cell["coordinate"] == "A3")
        self.assertEqual(formula["formula_text"], "=SUM(A1:A2)")
        self.assertIn("cached_value", formula)
        self.assertEqual(sheet["merged_ranges"], ["B1:C1"])
        self.assertEqual(len(sheet["charts"]), 1)

    def test_legacy_is_metadata_only_and_writer_compatible(self):
        path = self.root / "old.xls"
        path.write_bytes(b"not a readable legacy workbook")
        classification = classify_document(path)
        self.assertFalse(classification.readable)
        self.assertEqual(classification.kind, "legacy")
        report = inspect_file(path)
        self.assertIn("legacy_format_unsupported", report["risks"])
        self.assertFalse(report["macros"])

    def test_ooxml_external_relationship_and_macro_flags(self):
        path = self.root / "risky.xlsx"
        rels = b'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship TargetMode="External" Target="http://example.invalid" /></Relationships>'
        with zipfile.ZipFile(path, "w") as package:
            package.writestr("[Content_Types].xml", "<Types/>")
            package.writestr("xl/_rels/workbook.xml.rels", rels)
            package.writestr("xl/vbaProject.bin", b"macro")
        c = classify_document(path)
        self.assertTrue(c.macro_enabled)
        self.assertTrue(c.external_relationships)
        self.assertIn("external_relationship", c.suspicious_features)
        self.assertTrue(inspect_file(path)["external_links"])

    def test_ocr_result_is_explicit_and_non_mutating_contract(self):
        result = OCRResult("recognized", source="scan-1", confidence=0.9)
        self.assertEqual(result.source, "scan-1")
        self.assertEqual(result.confidence, 0.9)


if __name__ == "__main__":
    unittest.main()
