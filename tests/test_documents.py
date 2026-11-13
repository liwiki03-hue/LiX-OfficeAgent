import base64
import json
import tempfile
import unittest
import zipfile
from pathlib import Path

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference

try:
    from PIL import Image as PILImage
except ImportError:
    PILImage = None

from office_agent.documents import (
    OCRResult,
    UnsupportedDocumentError,
    batch_document_report,
    ocr_docx_image,
    classify_document,
    inspect_docx,
    inspect_file,
    inspect_xlsx,
    locate_word_cell,
    read_excel_structure,
    read_word_structure,
    read_word_table_structure,
    read_word_text,
    search_docx,
    search_document,
    summarize_word_document,
)
from office_agent.writes import Editor


_PIXEL_PNG = bytes.fromhex(
    "89504e470d0a1a0a0000000d494844520000000200000002080600000072b60d24"
    "0000001549444154789c636460f8ff9f8181818109448030001d190202b16791f6"
    "0000000049454e44ae426082")
_DRAWING = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"'
    ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
    ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
    '<xdr:oneCellAnchor><xdr:from><xdr:col>{column}</xdr:col><xdr:colOff>0</xdr:colOff>'
    '<xdr:row>{row}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>'
    '<xdr:ext cx="190500" cy="190500"/>'
    '<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="2" name="Picture 1"/><xdr:cNvPicPr/></xdr:nvPicPr>'
    '<xdr:blipFill><a:blip r:embed="rId1"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>'
    '<xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="190500" cy="190500"/></a:xfrm>'
    '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic>'
    '<xdr:clientData/></xdr:oneCellAnchor></xdr:wsDr>').replace("'", '"')


def _inject_picture(path, *, sheet_part, column, row):
    """Add a real drawing part so the image path can be tested without Pillow.

    Pillow is only needed to decode image bytes; injecting the OOXML directly keeps
    the image-detection path covered on an interpreter without it.
    """
    with zipfile.ZipFile(path) as archive:
        parts = {name: archive.read(name) for name in archive.namelist()}
    directory, _, name = sheet_part.rpartition("/")
    rels_path = f"{directory}/_rels/{name}.rels"
    sheet_rels = parts.get(rels_path, (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '</Relationships>').encode())
    parts[rels_path] = sheet_rels.replace(
        b"</Relationships>",
        b'<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/'
        b'relationships/drawing" Target="../drawings/drawing1.xml"/></Relationships>')
    parts["xl/drawings/drawing1.xml"] = _DRAWING.format(column=column, row=row).encode()
    parts["xl/drawings/_rels/drawing1.xml.rels"] = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/'
        'relationships/image" Target="../media/image1.png"/></Relationships>').encode()
    parts["xl/media/image1.png"] = _PIXEL_PNG
    parts["[Content_Types].xml"] = parts["[Content_Types].xml"].replace(
        b"</Types>",
        b'<Default Extension="png" ContentType="image/png"/>'
        b'<Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.'
        b'openxmlformats-officedocument.drawing+xml"/></Types>')
    sheet = parts[sheet_part]
    # openpyxl's saved worksheet omits the relationship namespace that <drawing r:id> needs.
    if b"xmlns:r=" not in sheet[:400]:
        sheet = sheet.replace(
            b"<worksheet ",
            b'<worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/'
            b'relationships" ', 1)
    parts[sheet_part] = sheet.replace(b"</worksheet>", b'<drawing r:id="rId1"/></worksheet>')
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as archive:
        for part, data in parts.items():
            archive.writestr(part, data)


_COMMENTS_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
    '<w:comment w:id="{id}" w:author="{author}" w:date="2024-01-01T00:00:00Z" w:initials="R">'
    '<w:p><w:r><w:t xml:space="preserve">{text}</w:t></w:r></w:p></w:comment></w:comments>')


def _add_comment(path, *, author, text, comment_id="1"):
    """Inject a real comments part; python-docx has no public API for adding comments."""
    with zipfile.ZipFile(path) as archive:
        parts = {name: archive.read(name) for name in archive.namelist()}
    parts["word/comments.xml"] = _COMMENTS_XML.format(
        id=comment_id, author=author, text=text).encode("utf-8")
    rels_path = "word/_rels/document.xml.rels"
    rels = parts[rels_path]
    parts[rels_path] = rels.replace(
        b"</Relationships>",
        b'<Relationship Id="rIdComments" Type="http://schemas.openxmlformats.org/officeDocument/'
        b'2006/relationships/comments" Target="comments.xml"/></Relationships>')
    parts["[Content_Types].xml"] = parts["[Content_Types].xml"].replace(
        b"</Types>",
        b'<Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-'
        b'officedocument.wordprocessingml.comments+xml"/></Types>')
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as archive:
        for part, data in parts.items():
            archive.writestr(part, data)


def _add_revision(paragraph, *, kind, text, revision_id, author="Reviewer"):
    """Inject a real tracked change; python-docx has no public API for revisions."""
    element = OxmlElement(f"w:{kind}")
    element.set(qn("w:id"), revision_id)
    element.set(qn("w:author"), author)
    element.set(qn("w:date"), "2024-05-01T00:00:00Z")
    run = OxmlElement("w:r")
    node = OxmlElement("w:t" if kind == "ins" else "w:delText")
    node.text = text
    run.append(node)
    element.append(run)
    paragraph._p.append(element)
    return element


def _add_content_control(document, *, tag, text, control="text"):
    """Append a block-level content control; python-docx has no public API for w:sdt."""
    structured = OxmlElement("w:sdt")
    properties = OxmlElement("w:sdtPr")
    tag_element = OxmlElement("w:tag")
    tag_element.set(qn("w:val"), tag)
    properties.append(tag_element)
    properties.append(OxmlElement(f"w:{control}"))
    structured.append(properties)
    content = OxmlElement("w:sdtContent")
    paragraph, run, node = OxmlElement("w:p"), OxmlElement("w:r"), OxmlElement("w:t")
    node.text = text
    run.append(node)
    paragraph.append(run)
    content.append(paragraph)
    structured.append(content)
    document.element.body.append(structured)
    return structured


class DocumentInspectionTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.root = Path(self.tmp.name)

    def tearDown(self):
        self.tmp.cleanup()

    def _pixel(self):
        image = self.root / "pixel.png"
        if not image.exists():
            image.write_bytes(base64.b64decode(
                "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/lZkAAAAASUVORK5CYII="))
        return image

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

    def test_dde_marker_requires_a_word_field_instruction(self):
        path = self.root / "ordinary.docx"
        doc = Document()
        doc.add_paragraph("A note mentioning DDE is not a field instruction")
        doc.save(path)
        self.assertNotIn("dde_marker", classify_document(path).suspicious_features)

        field = OxmlElement("w:fldSimple")
        field.set(qn("w:instr"), "DDEAUTO server topic")
        doc.paragraphs[0]._p.append(field)
        doc.save(path)
        self.assertIn("dde_marker", classify_document(path).suspicious_features)

    def test_docx_image_reports_relationship_position_and_placement(self):
        image = self.root / "pixel.png"
        image.write_bytes(base64.b64decode("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/lZkAAAAASUVORK5CYII="))
        path = self.root / "images.docx"
        doc = Document()
        doc.add_paragraph("before").add_run().add_picture(str(image))
        cell = doc.add_table(rows=1, cols=1).cell(0, 0)
        cell.paragraphs[0].add_run().add_picture(str(image))
        doc.save(path)

        report = inspect_docx(path)
        self.assertEqual(len(report["images"]), 2)
        locations = [item["position"] for item in report["images"]]
        self.assertCountEqual(locations, [
            {"type": "paragraph", "index": 0},
            {"type": "table_cell", "table": 0, "row": 0, "column": 0},
        ])
        # python-docx reuses the relationship for identical image bytes, but each
        # drawing still needs its own image entry and position.
        self.assertEqual(len({item["relationship"] for item in report["images"]}), 1)
        self.assertTrue(all(item["relationship"].startswith("rId") for item in report["images"]))
        self.assertEqual(len({item["target_part"] for item in report["images"]}), 1)
        self.assertTrue(all(item["source_part"] == "/word/document.xml" for item in report["images"]))
        self.assertTrue(all(item["placement"] == "inline" for item in report["images"]))
        self.assertEqual(report["image_placements"]["inline"], 2)

    def test_xlsx_image_position_is_read_from_the_package_without_pillow(self):
        """openpyxl silently drops images when Pillow is missing, so image presence and
        anchor position must come from the drawing part, not from ``ws._images``."""
        path = self.root / "image.xlsx"
        workbook = Workbook()
        worksheet = workbook.active
        worksheet.title = "Images"
        worksheet["A1"] = "hello"
        workbook.save(path)
        _inject_picture(path, sheet_part="xl/worksheets/sheet1.xml", column=2, row=3)

        sheet = inspect_xlsx(path)["sheets"][0]
        self.assertEqual([(item["type"], item["position"]) for item in sheet["images"]],
                         [("picture", {"type": "cell", "coordinate": "C4", "row": 4, "column": 3})])
        self.assertEqual(sheet["images"][0]["anchor"], "oneCellAnchor")
        self.assertEqual(sheet["images"][0]["target_part"], "xl/media/image1.png")
        # Content was not decoded here, and the report says so instead of implying absence.
        self.assertIs(sheet["images"][0]["bytes_decoded"], PILImage is not None)

    @unittest.skipUnless(PILImage is not None, "Pillow is not installed")
    def test_pillow_backed_image_metadata_still_matches_the_package_anchor(self):
        from openpyxl.drawing.image import Image as XLImage
        image = self.root / "pixel.png"
        PILImage.new("RGB", (2, 2), (255, 0, 0)).save(image, format="PNG")
        path = self.root / "image-pillow.xlsx"
        wb = Workbook()
        ws = wb.active
        ws.title = "Images"
        ws.add_image(XLImage(str(image)), "C4")
        wb.save(path)

        sheet = inspect_xlsx(path)["sheets"][0]
        self.assertEqual([item["position"] for item in sheet["images"]], [
            {"type": "cell", "coordinate": "C4", "row": 4, "column": 3}])
        self.assertTrue(sheet["images"][0]["bytes_decoded"])

    def test_sparse_far_cell_inspection_avoids_dense_region(self):
        path = self.root / "sparse.xlsx"
        workbook = Workbook()
        workbook.active["XFD1048576"] = "end"
        workbook.save(path)
        sheet = inspect_xlsx(path)["sheets"][0]
        self.assertEqual(sheet["cells"][0]["coordinate"], "XFD1048576")
        self.assertFalse(sheet["cells_truncated"])
        self.assertEqual(search_document(path, "end")[0]["coordinate"], "XFD1048576")

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

    def test_search_document_xlsx_streams_formula_and_value_matches_with_cap(self):
        path = self.root / "search.xlsx"
        wb = Workbook()
        ws = wb.active
        ws.title = "Data"
        ws["A1"] = "needle literal"
        ws["A2"] = "=CONCAT(\"needle\",\" formula\")"
        ws["A3"] = "needle again"
        wb.save(path)

        hits = search_document(path, "needle", max_results=2)
        self.assertEqual(len(hits), 2)
        self.assertEqual(hits[0]["sheet"], "Data")
        self.assertEqual(hits[0]["matched_field"], "value")
        self.assertEqual(hits[1]["matched_field"], "formula_text")
        self.assertEqual(hits[1]["formula_text"], '=CONCAT("needle"," formula")')
        self.assertIn("cached_value", hits[1])

    def test_search_docx_returns_bounded_snippets_and_locations_with_cap(self):
        path = self.root / "search.docx"
        doc = Document()
        doc.add_paragraph("needle " * 20)
        doc.add_table(rows=1, cols=1).cell(0, 0).text = "needle in table"
        doc.save(path)

        hits = search_docx(path, "needle", chunk_size=12, max_results=2)
        self.assertEqual(len(hits), 2)
        self.assertTrue(all(len(hit["text"]) <= 12 for hit in hits))
        self.assertEqual(hits[0]["location"]["type"], "paragraph")
        self.assertEqual(hits[1]["location"]["type"], "paragraph")

    def test_plain_zip_missing_office_parts_is_not_readable(self):
        path = self.root / "not-a-document.docx"
        with zipfile.ZipFile(path, "w") as archive:
            archive.writestr("arbitrary.txt", "hello")
        classification = classify_document(path)
        self.assertFalse(classification.readable)
        self.assertIn("missing_ooxml_core_part", classification.suspicious_features)

    def test_mislabeled_plaintext_ooxml_is_not_readable(self):
        for suffix in (".docx", ".xlsx"):
            path = self.root / ("invalid" + suffix)
            path.write_bytes(b"not a ZIP package")
            classification = classify_document(path)
            self.assertFalse(classification.readable)
            self.assertIn("invalid_zip", classification.suspicious_features)
            self.assertIn("invalid_zip", inspect_file(path)["risks"])

    def test_legacy_is_metadata_only_and_writer_compatible(self):
        path = self.root / "old.xls"
        path.write_bytes(b"not a readable legacy workbook")
        classification = classify_document(path)
        self.assertFalse(classification.readable)
        self.assertEqual(classification.kind, "legacy")
        report = inspect_file(path)
        self.assertIn("legacy_format_unsupported", report["risks"])
        self.assertIsNone(report["macros"])  # Legacy macro status is unknown without a reader.

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

    def test_word_structure_reads_headers_footers_comments_and_image_positions(self):
        """These were present in inspect_docx but had no model-facing bounded tool."""
        path = self.root / "structure.docx"
        doc = Document()
        doc.add_paragraph("Body paragraph that structure reads must not return")
        doc.sections[0].header.paragraphs[0].text = "Confidential header"
        doc.sections[0].footer.paragraphs[0].text = "Page footer"
        doc.add_picture(str(self._pixel()))
        doc.save(path)
        _add_comment(path, author="Reviewer", text="Please double-check this figure")

        report = read_word_structure(path, limit=5)
        self.assertEqual(report["sections"][0]["header"][0]["text"], "Confidential header")
        self.assertEqual(report["sections"][0]["footer"][0]["text"], "Page footer")
        self.assertEqual([c["author"] for c in report["comments"]], ["Reviewer"])
        self.assertEqual(report["comments"][0]["text"], "Please double-check this figure")
        self.assertEqual([i["position"] for i in report["images"]], [{"type": "paragraph", "index": 1}])
        self.assertEqual(report["images"][0]["placement"], "inline")
        self.assertEqual(report["image_placements"], {"inline": 1, "anchor": 0})
        # Body text is not materialised here; read_word_region is the tool for that.
        self.assertNotIn("Body paragraph", json.dumps(report, ensure_ascii=False))
        self.assertFalse(any(report[key] for key in
                             ("sections_truncated", "comments_truncated", "images_truncated")))

    def test_word_structure_is_bounded_and_reports_truncation(self):
        path = self.root / "many.docx"
        doc = Document()
        header = doc.sections[0].header
        header.paragraphs[0].text = "first"
        for index in range(6):
            header.add_paragraph(f"header line {index} " + "x" * 3000)
        for _ in range(4):
            doc.add_picture(str(self._pixel()))
        for _ in range(3):
            doc.add_section()
        doc.save(path)
        _add_comment(path, author="A", text="c" * 5000)

        report = read_word_structure(path, limit=3)
        self.assertEqual(len(report["sections"]), 3)
        self.assertTrue(report["sections_truncated"])  # four sections, three allowed
        self.assertTrue(report["images_truncated"])
        self.assertEqual(len(report["images"]), 3)
        self.assertEqual(len(report["sections"][0]["header"]), 3)
        self.assertTrue(report["sections"][0]["header_truncated"])
        # Oversized text is clipped rather than returned whole.
        self.assertTrue(report["sections"][0]["header"][1]["text_truncated"])
        self.assertEqual(len(report["sections"][0]["header"][1]["text"]), 1024)
        self.assertTrue(report["comments"][0]["text_truncated"])
        self.assertLessEqual(len(report["comments"][0]["text"]), 1024)
        for bad in (0, 51, -1):
            with self.assertRaises(ValueError):
                read_word_structure(path, limit=bad)

    def test_word_structure_reports_tracked_changes_and_content_controls(self):
        """Revisions and controls were previously invisible: nothing reported them at all."""
        path = self.root / "revisions.docx"
        doc = Document()
        inserted = doc.add_paragraph("kept text ")
        _add_revision(inserted, kind="ins", text="added by reviewer", revision_id="7", author="Alice")
        deleted = doc.add_paragraph("kept text ")
        _add_revision(deleted, kind="del", text="removed by reviewer", revision_id="8", author="Bob")
        doc.add_paragraph("clean paragraph")
        _add_content_control(doc, tag="ClientName", text="Acme Ltd")
        doc.save(path)

        report = read_word_structure(path, limit=10)
        changes = report["tracked_changes"]
        self.assertEqual([item["text"] for item in changes["insertions"]], ["added by reviewer"])
        self.assertEqual([item["text"] for item in changes["deletions"]], ["removed by reviewer"])
        self.assertEqual(changes["insertions"][0]["author"], "Alice")
        self.assertEqual(changes["insertions"][0]["id"], "7")
        self.assertEqual(changes["deletions"][0]["author"], "Bob")
        self.assertEqual(changes["insertion_count"], 1)
        self.assertEqual(changes["deletion_count"], 1)
        self.assertEqual(changes["move_mark_count"], 0)
        self.assertFalse(changes["insertions_truncated"] or changes["deletions_truncated"])
        self.assertTrue(changes["complete"])
        self.assertEqual(changes["parts_scanned"], ["word/document.xml"])
        self.assertEqual(report["content_controls"], [{"level": "block", "tag": "ClientName",
                                                       "alias": None, "control_type": "text",
                                                       "text": "Acme Ltd", "text_truncated": False}])
        self.assertEqual(report["content_control_count"], 1)
        self.assertFalse(report["content_controls_truncated"])

    def test_word_structure_counts_revisions_beyond_the_limit_instead_of_hiding_them(self):
        path = self.root / "many-revisions.docx"
        doc = Document()
        for index in range(5):
            paragraph = doc.add_paragraph("text ")
            _add_revision(paragraph, kind="ins", text="i" * 3000, revision_id=str(index))
        for index in range(3):
            _add_content_control(doc, tag=f"Field{index}", text="v" * 3000)
        doc.save(path)

        report = read_word_structure(path, limit=2)
        changes = report["tracked_changes"]
        self.assertEqual(len(changes["insertions"]), 2)
        self.assertEqual(changes["insertion_count"], 5)   # the count is never truncated
        self.assertTrue(changes["insertions_truncated"])
        self.assertEqual(changes["deletion_count"], 0)
        # Oversized revision text is clipped and marked, exactly like the other structure reads.
        self.assertTrue(changes["insertions"][0]["text_truncated"])
        self.assertEqual(len(changes["insertions"][0]["text"]), 1024)
        self.assertEqual(len(report["content_controls"]), 2)
        self.assertEqual(report["content_control_count"], 3)
        self.assertTrue(report["content_controls_truncated"])
        self.assertTrue(report["content_controls"][0]["text_truncated"])

    def test_word_revision_report_fails_closed_on_an_unparseable_part(self):
        """An unreadable part must refuse the read, never return "no revisions found"."""
        from office_agent.documents import _word_revision_report
        path = self.root / "broken-revisions.docx"
        with zipfile.ZipFile(path, "w") as archive:
            archive.writestr("word/document.xml", "<w:document>")  # truncated XML
        with zipfile.ZipFile(path) as archive:
            with self.assertRaises(UnsupportedDocumentError):
                _word_revision_report(archive, limit=5)

    def test_excel_structure_reports_named_ranges_formatting_hidden_axes_and_charts(self):
        from openpyxl.formatting.rule import CellIsRule
        from openpyxl.workbook.defined_name import DefinedName
        path = self.root / "structure.xlsx"
        wb = Workbook(); ws = wb.active; ws.title = "Data"
        ws["A1"] = 1
        ws["A2"] = 2
        ws.defined_names.add(DefinedName("LocalTotal", attr_text="Data!$A$1:$A$2"))
        wb.defined_names.add(DefinedName("Rate", attr_text='"5%"'))
        wb.defined_names.add(DefinedName("Computed", attr_text="SUM(Data!$A$1:$A$2)"))
        ws.conditional_formatting.add("A1:A2", CellIsRule(operator="greaterThan", formula=["1"]))
        ws.row_dimensions[3].hidden = True
        ws.row_dimensions[4].hidden = True
        ws.column_dimensions["C"].hidden = True
        ws.merge_cells("E1:F1")
        chart = BarChart()
        chart.title = "Revenue"
        chart.add_data(Reference(ws, min_col=1, min_row=1, max_row=2), titles_from_data=False)
        ws.add_chart(chart, "D5")
        wb.save(path)

        report = read_excel_structure(path, limit=20)
        self.assertTrue(report["readable"])
        self.assertEqual(report["sheet_titles"], ["Data"])
        names = {entry["name"]: entry for entry in report["defined_names"]}
        self.assertEqual(names["LocalTotal"]["scope"], "Data")     # sheet-scoped name
        self.assertEqual(names["LocalTotal"]["kind"], "range")
        self.assertEqual(names["Rate"]["scope"], "workbook")
        self.assertEqual(names["Rate"]["kind"], "constant")
        self.assertEqual(names["Computed"]["kind"], "formula_or_unknown")  # never guessed at
        self.assertEqual(report["defined_name_count"], 3)
        self.assertFalse(report["defined_names_truncated"])

        sheet = report["sheets"][0]
        self.assertEqual([(run["first"], run["last"], run["count"]) for run in sheet["hidden_rows"]],
                         [(3, 4, 2)])
        self.assertEqual(sheet["hidden_row_count"], 2)
        self.assertEqual([run["label"] for run in sheet["hidden_columns"]], ["C"])
        self.assertEqual(sheet["hidden_column_count"], 1)
        self.assertEqual([entry["range"] for entry in sheet["conditional_formatting"]], ["A1:A2"])
        self.assertEqual(sheet["conditional_formatting"][0]["rules"][0]["type"], "cellIs")
        self.assertEqual(sheet["conditional_formatting"][0]["rules"][0]["formula"], ["1"])
        self.assertEqual(sheet["conditional_formatting_rule_count"], 1)
        self.assertEqual(sheet["merged_ranges"], ["E1:F1"])
        self.assertEqual(sheet["charts"][0]["title"], "Revenue")
        self.assertFalse(sheet["charts"][0]["title_unreadable"])
        self.assertEqual(sheet["charts"][0]["anchor"],
                         {"type": "cell", "coordinate": "D5", "row": 5, "column": 4})
        self.assertEqual(sheet["charts"][0]["series"],
                         [{"values": "'Data'!$A$1:$A$2", "categories": None}])
        self.assertEqual(report["chart_count"], 1)
        self.assertTrue(report["drawings_readable"])
        self.assertFalse(any(sheet[key] for key in ("hidden_rows_truncated", "hidden_columns_truncated",
                                                    "charts_truncated", "conditional_formatting_truncated",
                                                    "images_truncated", "merged_ranges_truncated")))

    def test_excel_structure_is_bounded_and_flags_every_truncated_list(self):
        from openpyxl.formatting.rule import FormulaRule
        from openpyxl.workbook.defined_name import DefinedName
        path = self.root / "structure-bounded.xlsx"
        wb = Workbook(); ws = wb.active; ws.title = "Data"
        ws["A1"] = 1
        ws["A2"] = 2
        wb.defined_names.add(DefinedName("One", attr_text="Data!$A$1"))
        wb.defined_names.add(DefinedName("Two", attr_text="Data!$A$2"))
        wb.defined_names.add(DefinedName("Three", attr_text="Data!$A$1:$A$2"))
        for row in range(1, 5):
            ws.conditional_formatting.add(f"A{row}:A{row}", FormulaRule(formula=[f"$A{row}>0"]))
        for row in (1, 3, 5):                      # three separate hidden blocks
            ws.row_dimensions[row].hidden = True
        for column in ("C", "E", "G"):             # three separate hidden blocks
            ws.column_dimensions[column].hidden = True
        for anchor in ("G1", "G20"):
            chart = BarChart()
            chart.add_data(Reference(ws, min_col=1, min_row=1, max_row=2), titles_from_data=False)
            ws.add_chart(chart, anchor)
        ws.merge_cells("J1:K1")
        ws.merge_cells("J3:K3")
        ws.merge_cells("J5:K5")
        wb.save(path)

        report = read_excel_structure(path, limit=2)
        self.assertTrue(report["defined_names_truncated"])
        self.assertEqual(report["defined_name_count"], 3)
        self.assertEqual(len(report["defined_names"]), 2)
        sheet = report["sheets"][0]
        self.assertEqual(len(sheet["hidden_rows"]), 2)
        self.assertEqual(sheet["hidden_row_count"], 3)          # the count stays complete
        self.assertTrue(sheet["hidden_rows_truncated"])
        self.assertEqual(sheet["hidden_column_count"], 3)
        self.assertTrue(sheet["hidden_columns_truncated"])
        self.assertEqual(len(sheet["conditional_formatting"]), 2)
        self.assertEqual(sheet["conditional_formatting_range_count"], 4)
        self.assertTrue(sheet["conditional_formatting_truncated"])
        self.assertEqual(len(sheet["charts"]), 2)
        self.assertEqual(sheet["chart_count"], 2)
        self.assertFalse(sheet["charts_truncated"])
        self.assertEqual(sheet["merged_range_count"], 3)
        self.assertTrue(sheet["merged_ranges_truncated"])
        for bad in (0, 101, -3):
            with self.assertRaises(ValueError):
                read_excel_structure(path, limit=bad)

    def test_excel_structure_says_when_totals_only_cover_the_detailed_sheets(self):
        path = self.root / "many-sheets.xlsx"
        wb = Workbook()
        for index in range(21):
            ws = wb.active if index == 0 else wb.create_sheet()
            ws.title = f"S{index}"
            ws["A1"] = index
        wb.save(path)

        report = read_excel_structure(path)
        self.assertEqual(report["sheet_count"], 21)
        self.assertEqual(len(report["sheets"]), 20)
        self.assertTrue(report["sheets_truncated"])
        # A total over 20 of 21 sheets must not look like a total over the workbook.
        self.assertTrue(report["counts_cover_detailed_sheets_only"])

    def test_excel_structure_refuses_unreadable_or_protected_input(self):
        path = self.root / "not-a-workbook.xlsx"
        with zipfile.ZipFile(path, "w") as archive:
            archive.writestr("arbitrary.txt", "hello")
        with self.assertRaises(UnsupportedDocumentError):
            read_excel_structure(path)
        plain = self.root / "plain.xlsx"
        plain.write_bytes(b"not a ZIP package")
        with self.assertRaises(UnsupportedDocumentError):
            read_excel_structure(plain)
        with self.assertRaises(UnsupportedDocumentError):
            read_excel_structure(self.root / "missing.xlsx")

    def _merged_table_doc(self, name="merged.docx"):
        """3x3 table: a gridSpan pair in row 0, a vMerge pair in column 0, one nested table."""
        path = self.root / name
        doc = Document()
        table = doc.add_table(rows=3, cols=3)
        table.style = "Table Grid"
        table.cell(0, 0).text = "span A"
        table.cell(0, 0).merge(table.cell(0, 1))     # w:gridSpan
        table.cell(0, 2).text = "span B"
        table.cell(1, 0).text = "vertical"
        table.cell(1, 0).merge(table.cell(2, 0))     # w:vMerge restart + continue
        table.cell(1, 1).text = "plain 1"
        table.cell(2, 1).text = "plain 2"
        nested = table.cell(2, 2).add_table(rows=1, cols=2)
        nested.cell(0, 0).text = "inner"
        doc.save(path)
        return path

    def test_word_table_structure_resolves_spans_vertical_merges_and_nesting(self):
        """Both merges were previously invisible: nothing reported gridSpan/vMerge at all."""
        path = self._merged_table_doc()
        report = read_word_table_structure(path, limit=10)

        self.assertEqual(report["grid_columns"], 3)
        self.assertEqual(report["row_count"], 3)
        self.assertTrue(report["merge_consistent"])
        self.assertEqual(report["merge_issues"], [])
        span_row = report["rows"][0]
        self.assertEqual(span_row["grid_covered"], 3)
        self.assertEqual([(cell["column"], cell["grid_span"], cell["last_column"])
                          for cell in span_row["cells"]], [(0, 2, 1), (2, 1, 2)])
        # Physical and logical coordinates stay separate: two w:tc cover three grid columns.
        self.assertEqual([cell["physical_index"] for cell in span_row["cells"]], [0, 1])
        self.assertEqual(span_row["cells"][0]["anchor"], {"row": 0, "column": 0})

        vertical = report["rows"][2]["cells"][0]
        self.assertEqual(vertical["vmerge"], "continue")
        self.assertEqual(vertical["anchor"], {"row": 1, "column": 0})   # the origin cell
        self.assertEqual(vertical["text"], "")
        self.assertEqual(report["rows"][1]["cells"][0]["vmerge"], "restart")
        self.assertEqual(report["continuations"],
                         [{"row": 2, "column": 0, "anchor": {"row": 1, "column": 0}, "grid_span": 1}])
        self.assertEqual(report["rows"][2]["cells"][2]["nested_table_count"], 1)
        self.assertEqual(report["nested_tables"],
                         [{"depth": 1, "path": [{"row": 2, "column": 2}], "rows": 1,
                           "grid_columns": 2, "merge_consistent": True}])
        self.assertFalse(report["depth_limit_reached"])
        self.assertEqual(report["table_index"], 0)
        self.assertEqual(report["table_count"], 1)

    def test_word_table_structure_reports_inconsistency_instead_of_inventing_a_merge(self):
        path = self.root / "broken-table.docx"
        doc = Document()
        table = doc.add_table(rows=2, cols=3)
        table.cell(0, 0).text = "a"
        table.cell(0, 1).text = "b"
        table.cell(0, 2).text = "c"
        table.cell(1, 0).text = "d"
        table.cell(1, 1).text = "e"
        table.cell(1, 2).text = "f"
        doc.save(path)
        # Craft three defects the report must name rather than smooth over: a vMerge
        # continuation with no restart above, a row that no longer covers the grid, and a
        # legacy hMerge element.
        document = Document(path)
        tbl = document.tables[0]._tbl
        row_one = tbl.findall(qn("w:tr"))[1]
        first_cell = row_one.findall(qn("w:tc"))[0]
        merge = OxmlElement("w:vMerge")
        first_cell.find(qn("w:tcPr")).append(merge)
        row_one.remove(row_one.findall(qn("w:tc"))[-1])          # coverage now 2 of 3
        legacy = OxmlElement("w:hMerge")
        first_cell.find(qn("w:tcPr")).append(legacy)
        document.save(path)

        report = read_word_table_structure(path, limit=10)
        self.assertFalse(report["merge_consistent"])
        issues = " | ".join(report["merge_issues"])
        self.assertIn("w:vMerge 延续格在上方找不到起点", issues)
        self.assertIn("与 w:tblGrid 的 3 列不一致", issues)
        self.assertIn("w:hMerge", issues)
        self.assertEqual(report["merge_issue_count"], len(report["merge_issues"]))
        self.assertFalse(report["merge_issues_truncated"])
        # An orphan continuation keeps its own coordinate instead of being given an origin.
        self.assertEqual(report["rows"][1]["cells"][0]["anchor"], {"row": 1, "column": 0})

    def test_word_table_structure_is_bounded_and_flags_truncation_and_depth(self):
        path = self.root / "deep-table.docx"
        doc = Document()
        table = doc.add_table(rows=5, cols=2)
        for index in range(5):
            table.cell(index, 0).text = f"r{index}"
        level = table.cell(0, 1).add_table(rows=1, cols=1)
        for _ in range(3):                                        # four levels of nesting
            level = level.cell(0, 0).add_table(rows=1, cols=1)
        doc.save(path)

        report = read_word_table_structure(path, limit=10)
        self.assertEqual(report["row_count"], 5)
        self.assertEqual(len(report["rows"]), 5)
        self.assertFalse(report["cells_truncated"])
        # Four levels exist; the report expands three and says the cap was reached.
        self.assertEqual([entry["depth"] for entry in report["nested_tables"]], [1, 2, 3])
        self.assertTrue(report["depth_limit_reached"])
        self.assertFalse(report["nested_tables_truncated"])
        self.assertEqual(report["max_nested_depth"], 3)

        bounded = read_word_table_structure(path, limit=2)
        self.assertEqual(bounded["row_count"], 5)
        self.assertEqual(len(bounded["rows"]), 2)                 # bounded, not silently short
        self.assertTrue(bounded["cells_truncated"])
        self.assertEqual(len(bounded["nested_tables"]), 2)
        self.assertTrue(bounded["nested_tables_truncated"])
        for bad in (0, 101, -1):
            with self.assertRaises(ValueError):
                read_word_table_structure(path, limit=bad)
        with self.assertRaises(ValueError):
            read_word_table_structure(path, table_index=-1)

    def test_word_table_structure_refuses_non_tables_and_out_of_range_index(self):
        docx_path = self.root / "no-tables.docx"
        document = Document()
        document.add_paragraph("no table here")
        document.save(docx_path)
        with self.assertRaisesRegex(ValueError, "out of bounds"):
            read_word_table_structure(docx_path)
        xlsx = self.root / "book.xlsx"
        workbook = Workbook()
        workbook.save(xlsx)
        with self.assertRaises(UnsupportedDocumentError):
            read_word_table_structure(xlsx)
        with self.assertRaises(UnsupportedDocumentError):
            read_word_table_structure(self.root / "missing.docx")

    def test_locate_word_cell_names_the_origin_for_covered_and_continued_cells(self):
        path = self._merged_table_doc("locate.docx")
        tbl = Document(path).tables[0]._tbl
        element, info = locate_word_cell(tbl, 0, 0)
        self.assertEqual(info["grid_span"], 2)
        self.assertEqual(info["text"], "span A")
        self.assertEqual(element.tag, qn("w:tc"))
        with self.assertRaisesRegex(ValueError, r"被 \(行 0, 列 0\) 的跨列合并覆盖"):
            locate_word_cell(tbl, 0, 1)
        with self.assertRaisesRegex(ValueError, r"文本由 \(行 1, 列 0\) 承载"):
            locate_word_cell(tbl, 2, 0)
        with self.assertRaisesRegex(ValueError, "超出表格范围"):
            locate_word_cell(tbl, 9, 0)
        with self.assertRaisesRegex(ValueError, "超出表格范围"):
            locate_word_cell(tbl, 0, 9)

    def test_batch_report_summarises_the_authorized_directory_and_writes_nothing(self):
        root = self.root / "batch"
        root.mkdir()
        doc = Document(); doc.add_paragraph("Alpha report"); doc.add_paragraph("needle here")
        doc.save(root / "a.docx")
        doc = Document(); doc.add_heading("Beta", level=1); doc.add_paragraph("needle again")
        doc.save(root / "b.docx")
        workbook = Workbook(); workbook.active["A1"] = 1; workbook.active["A2"] = "needle"
        workbook.save(root / "c.xlsx")
        (root / "old.doc").write_bytes(b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1legacy")
        (root / "broken.docx").write_bytes(b"not a zip package")
        (root / "notes.txt").write_text("unrelated", encoding="utf-8")
        before = {path.name: path.read_bytes() for path in root.iterdir()}

        space = Editor(root).space
        report = batch_document_report(root / "a.docx", validate=space.file, operation="summary")
        self.assertTrue(report["wrote_nothing"])
        self.assertEqual([entry["file"] for entry in report["files"]], ["a.docx", "b.docx", "c.xlsx"])
        self.assertEqual(report["file_count"], 3)
        self.assertEqual(report["error_count"], 0)
        self.assertFalse(report["truncated"])
        word = report["files"][0]
        self.assertEqual((word["kind"], word["paragraph_count"], word["character_count"]),
                         ("word", 2, 23))
        self.assertEqual(word["preview"], ["Alpha report needle here"])
        self.assertEqual(report["files"][1]["headings"], ["Beta"])
        self.assertEqual(report["files"][2]["sheet_count"], 1)
        self.assertEqual(report["files"][2]["sheets"][0]["populated_cells"], 2)
        reasons = {entry["file"]: entry["reason"] for entry in report["skipped"]}
        self.assertIn("遗留格式 .doc", reasons["old.doc"])
        self.assertIn("结构识别报风险", reasons["broken.docx"])
        self.assertNotIn("notes.txt", reasons)          # unrelated files are not batch inputs
        # Read-only means read-only: not one byte of the directory changed.
        self.assertEqual({path.name: path.read_bytes() for path in root.iterdir()}, before)

    def test_batch_report_validates_every_candidate_with_the_same_guard(self):
        root = self.root / "guarded"
        root.mkdir()
        for name in ("a.docx", "b.docx", "c.docx"):
            document = Document(); document.add_paragraph(name); document.save(root / name)
        seen: list[str] = []
        space = Editor(root).space

        def validate(raw):
            seen.append(Path(raw).name)
            if Path(raw).name == "b.docx":
                raise ValueError("file outside authorized directory")
            return space.file(raw)

        report = batch_document_report(root / "a.docx", validate=validate)
        self.assertEqual(sorted(seen), ["a.docx", "b.docx", "c.docx"])
        self.assertEqual([entry["file"] for entry in report["files"]], ["a.docx", "c.docx"])
        refused = [entry for entry in report["skipped"] if entry["file"] == "b.docx"]
        self.assertEqual(len(refused), 1)
        self.assertIn("路径护栏拒绝", refused[0]["reason"])
        self.assertIn("file outside authorized directory", refused[0]["reason"])

    def test_batch_report_skips_a_symlink_out_of_the_directory(self):
        root = self.root / "linked"
        root.mkdir()
        outside = self.root / "outside-secret.docx"
        document = Document(); document.add_paragraph("secret"); document.save(outside)
        document = Document(); document.add_paragraph("inside"); document.save(root / "inside.docx")
        link = root / "escape.docx"
        try:
            link.symlink_to(outside)
        except OSError as exc:                          # ordinary users cannot always symlink
            self.skipTest(f"symlink unavailable: {exc}")
        report = batch_document_report(root / "inside.docx", validate=Editor(root).space.file)
        self.assertEqual([entry["file"] for entry in report["files"]], ["inside.docx"])
        refused = [entry for entry in report["skipped"] if entry["file"] == "escape.docx"]
        self.assertEqual(len(refused), 1)
        self.assertIn("路径护栏拒绝", refused[0]["reason"])
        self.assertNotIn("secret", json.dumps(report, ensure_ascii=False))

    def test_batch_report_bounds_files_and_flags_truncation(self):
        root = self.root / "many"
        root.mkdir()
        for index in range(25):
            document = Document(); document.add_paragraph(f"file {index}")
            document.save(root / f"doc-{index:02d}.docx")
        report = batch_document_report(root / "doc-00.docx", validate=Editor(root).space.file,
                                       limit=20, per_file=5)
        self.assertEqual(report["file_count"], 20)
        self.assertTrue(report["truncated"])
        self.assertEqual(report["limit"], 20)
        with self.assertRaisesRegex(ValueError, "limit must be between 1 and 20"):
            batch_document_report(root / "doc-00.docx", validate=Editor(root).space.file, limit=21)
        with self.assertRaisesRegex(ValueError, "per_file must be between 1 and 20"):
            batch_document_report(root / "doc-00.docx", validate=Editor(root).space.file, per_file=0)
        with self.assertRaisesRegex(ValueError, "operation must be"):
            batch_document_report(root / "doc-00.docx", validate=Editor(root).space.file,
                                  operation="delete")
        with self.assertRaisesRegex(ValueError, "search requires a non-empty query"):
            batch_document_report(root / "doc-00.docx", validate=Editor(root).space.file,
                                  operation="search")

    def test_batch_exact_limit_ignores_unrelated_later_siblings(self):
        root = self.root / "exact-limit"
        root.mkdir()
        for index in range(20):
            doc = Document(); doc.add_paragraph(f"synthetic {index}")
            doc.save(root / f"doc-{index:02d}.docx")
        (root / "z-nested.docx").mkdir()
        (root / "z-nested.docx" / "extra.docx").write_bytes(b"out of scope")
        (root / "z-notes.txt").write_text("unrelated", encoding="utf-8")
        report = batch_document_report(root / "doc-00.docx",
                                       validate=Editor(root).space.file, operation="inspect")
        self.assertEqual(report["file_count"], 20)
        self.assertFalse(report["truncated"])
        self.assertEqual(report["skipped_count"], 0)
        self.assertTrue(report["wrote_nothing"])

    def test_batch_report_isolates_one_unreadable_file_and_keeps_going(self):
        root = self.root / "isolated"
        root.mkdir()
        document = Document(); document.add_paragraph("fine"); document.save(root / "fine.docx")
        # A structurally valid package that passes the safety inspector but cannot be parsed.
        with zipfile.ZipFile(root / "bad.docx", "w") as archive:
            archive.writestr("[Content_Types].xml", "<Types/>")
            archive.writestr("word/document.xml", "<w:document>")
        report = batch_document_report(root / "fine.docx", validate=Editor(root).space.file)
        self.assertEqual([entry["file"] for entry in report["files"]], ["fine.docx"])
        self.assertEqual(report["error_count"], 1)
        self.assertEqual(report["errors"][0]["file"], "bad.docx")
        self.assertNotIn("bad.docx", [entry["file"] for entry in report["files"]])

    def test_batch_search_reports_bounded_hits_per_file(self):
        root = self.root / "search"
        root.mkdir()
        for name in ("one.docx", "two.docx"):
            document = Document()
            for _ in range(5):
                document.add_paragraph("needle in a paragraph")
            document.save(root / name)
        report = batch_document_report(root / "one.docx", validate=Editor(root).space.file,
                                       operation="search", query="needle", per_file=2)
        self.assertEqual(report["file_count"], 2)
        for entry in report["files"]:
            self.assertEqual(entry["hit_count"], 2)
            self.assertTrue(entry["hits_truncated"])
        report = batch_document_report(root / "one.docx", validate=Editor(root).space.file,
                                       operation="inspect")
        self.assertTrue(all(entry["readable"] for entry in report["files"]))
        self.assertEqual(report["files"][0]["risks"], [])

    def test_batch_report_excludes_own_output_copies_and_office_locks(self):
        root = self.root / "outputs"
        root.mkdir()
        document = Document(); document.add_paragraph("source"); document.save(root / "src.docx")
        (root / "src.agent-1234abcd.docx").write_bytes((root / "src.docx").read_bytes())
        (root / "~$src.docx").write_bytes(b"lock")
        report = batch_document_report(root / "src.docx", validate=Editor(root).space.file)
        self.assertEqual([entry["file"] for entry in report["files"]], ["src.docx"])
        reasons = {entry["file"]: entry["reason"] for entry in report["skipped"]}
        self.assertIn("输出副本", reasons["src.agent-1234abcd.docx"])
        self.assertNotIn("~$src.docx", reasons)          # Office lock files are not inputs at all

    def test_word_summary_maps_chunks_and_outline_within_bounds(self):
        """Reading every paragraph via read_word_region would need far more than the
        15-decision budget, so the summary path must be one bounded retrieval step."""
        path = self.root / "report.docx"
        doc = Document()
        doc.add_paragraph("Report title", style="Heading 1")
        for index in range(60):
            doc.add_paragraph(f"Body paragraph {index} " + "detail " * 10)
        doc.add_paragraph("Second section", style="Heading 2")
        for index in range(20):
            doc.add_paragraph(f"More body {index} " + "detail " * 10)
        doc.save(path)

        report = summarize_word_document(path, chunk_size=20, max_chunks=5, preview_chars=100)
        self.assertEqual(report["paragraph_count"], 82)
        self.assertEqual([h["text"] for h in report["headings"]], ["Report title", "Second section"])
        self.assertEqual([h["style"] for h in report["headings"]], ["Heading 1", "Heading 2"])
        self.assertEqual(report["total_chunks"], 5)
        self.assertEqual(len(report["chunks"]), 5)
        self.assertEqual(report["chunks"][0]["start_index"], 0)
        self.assertEqual(report["chunks"][0]["end_index"], 19)
        self.assertLessEqual(len(report["chunks"][0]["preview"]), 100)
        self.assertIn("Report title", report["chunks"][0]["preview"])
        self.assertTrue(all(chunk["characters"] > 0 for chunk in report["chunks"]))
        # The returned scaffold stays small even though the document does not.
        self.assertLess(len(json.dumps(report, ensure_ascii=False)), 5000)

    def test_word_summary_omits_excess_chunks_and_flags_it(self):
        path = self.root / "long.docx"
        doc = Document()
        for index in range(200):
            doc.add_paragraph(f"Line {index}")
        doc.save(path)
        report = summarize_word_document(path, chunk_size=10, max_chunks=4)
        self.assertEqual(report["total_chunks"], 20)
        self.assertEqual(len(report["chunks"]), 4)
        self.assertEqual(report["chunks_omitted"], 16)
        for bad in ({"chunk_size": 0}, {"max_chunks": 0}, {"preview_chars": 0},
                    {"chunk_size": 201}, {"max_chunks": 101}, {"preview_chars": 1001}):
            with self.assertRaises(ValueError):
                summarize_word_document(path, **bad)

    def test_word_summary_caps_a_long_heading_outline(self):
        path = self.root / "headings.docx"
        doc = Document()
        for index in range(120):
            doc.add_paragraph(f"Heading {index}", style="Heading 1")
        doc.save(path)
        report = summarize_word_document(path, chunk_size=50, max_chunks=4)
        self.assertEqual(len(report["headings"]), 100)
        self.assertTrue(report["headings_truncated"])
        # Truncation is reported, never silently presented as the whole outline.
        self.assertEqual(report["headings"][-1]["text"], "Heading 99")

    def test_word_summary_reports_table_shapes_so_table_content_is_not_missed(self):
        """A document whose content lives in a table looks empty through paragraphs alone,
        which is exactly what happened to the user's 36-character speech draft."""
        path = self.root / "table-heavy.docx"
        doc = Document()
        doc.add_paragraph("Intro line")
        table = doc.add_table(rows=6, cols=9)
        table.cell(0, 0).text = "the actual content lives here"
        doc.save(path)

        report = summarize_word_document(path)
        self.assertEqual(report["paragraph_count"], 1)
        self.assertEqual(report["table_count"], 1)
        self.assertEqual(report["tables"], [{"index": 0, "rows": 6, "columns": 9}])
        self.assertFalse(report["tables_truncated"])
        self.assertIn("read_word_table_region", report["note"])
        # Shapes only: no cell text is pulled into the summary.
        self.assertNotIn("the actual content lives here", json.dumps(report, ensure_ascii=False))

    def test_word_text_reads_body_and_tables_in_document_order(self):
        """The user's speech draft had 36 characters of body text and all its content in a
        table, so a "show me the text" request must not stop at the body."""
        path = self.root / "ordered.docx"
        doc = Document()
        doc.add_paragraph("Opening line")
        table = doc.add_table(rows=2, cols=2)
        table.cell(0, 0).text = "table cell A"
        table.cell(0, 1).text = "table cell B"
        doc.add_paragraph("Closing line")
        doc.save(path)

        report = read_word_text(path)
        self.assertEqual([b["kind"] for b in report["blocks"]],
                         ["paragraph", "table", "paragraph"])
        self.assertEqual(report["blocks"][0]["text"], "Opening line")
        self.assertEqual(report["blocks"][1]["rows"][0], ["table cell A", "table cell B"])
        self.assertEqual(report["blocks"][1]["row_count"], 2)
        self.assertEqual(report["blocks"][2]["text"], "Closing line")
        self.assertEqual(report["total_blocks"], 3)
        self.assertFalse(report["has_more"])

    def test_word_text_never_skips_a_block_that_exceeds_the_budget(self):
        path = self.root / "oversized.docx"
        doc = Document()
        doc.add_paragraph("short")
        table = doc.add_table(rows=6, cols=2)
        for row in table.rows:
            for cell in row.cells:
                cell.text = "x" * 500
        doc.save(path)

        report = read_word_text(path, max_chars=600)
        kinds = [b["kind"] for b in report["blocks"]]
        # The table must appear even though it cannot fit, marked as truncated.
        self.assertIn("table", kinds)
        table_block = next(b for b in report["blocks"] if b["kind"] == "table")
        self.assertTrue(table_block["rows_truncated"])
        self.assertLess(len(table_block["rows"]), 6)
        # 600 is below the floor, so the applied budget is the floor.
        self.assertEqual(report["effective_max_chars"], 1000)
        self.assertLessEqual(report["returned_chars"], report["effective_max_chars"])
        self.assertTrue(report["has_more"])

    def test_word_text_paginates_with_next_offset(self):
        path = self.root / "paged.docx"
        doc = Document()
        for index in range(6):
            doc.add_paragraph(f"line {index} " + "y" * 100)
        doc.save(path)
        seen, offset, rounds = [], 0, 0
        while rounds < 10:
            report = read_word_text(path, offset=offset, limit=50, max_chars=250)
            seen.extend(b["text"] for b in report["blocks"] if b["kind"] == "paragraph")
            offset = report["next_offset"]
            rounds += 1
            if not report["has_more"]:
                break
        self.assertEqual(len(seen), 6)
        self.assertEqual(seen[0][:6], "line 0")
        self.assertEqual(seen[-1][:6], "line 5")
        self.assertFalse(report["has_more"])
        for bad in ({"offset": -1}, {"limit": 0}, {"limit": 51}, {"max_chars": 100},
                    {"max_chars": 20001}):
            with self.assertRaises(ValueError):
                read_word_text(path, **bad)

    def test_word_text_raises_a_uselessly_small_request_to_a_working_floor(self):
        """Small models send the minimum allowed numbers; a 66-character read answers a
        "show me the text" request uselessly, so a floor is applied and reported."""
        path = self.root / "floored.docx"
        doc = Document()
        doc.add_paragraph("first line")
        table = doc.add_table(rows=2, cols=2)
        table.cell(0, 0).text = "content in a table"
        doc.save(path)

        report = read_word_text(path, limit=1, max_chars=200)
        self.assertEqual(report["limit"], 1)
        self.assertEqual(report["max_chars"], 200)
        self.assertEqual(report["effective_limit"], 5)
        self.assertEqual(report["effective_max_chars"], 1000)
        # The table is therefore reached even though the caller asked for one block.
        self.assertIn("table", [b["kind"] for b in report["blocks"]])
        self.assertIn("content in a table", json.dumps(report, ensure_ascii=False))
        # A larger explicit request is honoured as-is.
        larger = read_word_text(path, limit=30, max_chars=5000)
        self.assertEqual((larger["effective_limit"], larger["effective_max_chars"]), (30, 5000))

    def test_word_text_refuses_non_docx(self):
        xlsx = self.root / "notword.xlsx"
        Workbook().save(xlsx)
        with self.assertRaises(Exception):
            read_word_text(xlsx)

    def test_word_summary_refuses_non_docx(self):
        xlsx = self.root / "notword.xlsx"
        Workbook().save(xlsx)
        with self.assertRaises(Exception):
            summarize_word_document(xlsx)

    def test_word_structure_refuses_non_docx_and_unreadable_input(self):
        xlsx = self.root / "notword.xlsx"
        Workbook().save(xlsx)
        with self.assertRaises(Exception):
            read_word_structure(xlsx)
        with self.assertRaises(Exception):
            read_word_structure(self.root / "missing.docx")

    def test_ocr_result_is_explicit_and_non_mutating_contract(self):
        result = OCRResult("recognized", source="scan-1", confidence=0.9)
        self.assertEqual(result.source, "scan-1")
        self.assertEqual(result.confidence, 0.9)
        image = self.root / "pixel.png"
        image.write_bytes(base64.b64decode("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/lZkAAAAASUVORK5CYII="))
        path = self.root / "ocr.docx"
        doc = Document(); doc.add_paragraph("original text"); doc.add_picture(str(image)); doc.save(path)
        before = path.read_bytes()
        class FakeOCR:
            def recognize(self, image, *, source):
                return OCRResult("image text", source=source, confidence=0.87)
        report = inspect_docx(path, ocr=FakeOCR())
        self.assertEqual(report["ocr"][0]["confidence"], 0.87)
        self.assertIn("relationship=", report["ocr"][0]["source"])
        self.assertNotIn("image text", report["text"])
        rel = report["images"][0]["relationship"]
        single = ocr_docx_image(path, rel, FakeOCR())
        self.assertEqual(single["confidence"], 0.87)
        self.assertEqual(single["source"], f"{path}#relationship={rel}")
        with self.assertRaises(ValueError):
            ocr_docx_image(path, "rId999999", FakeOCR())
        class OversizedOCR:
            def recognize(self, image, *, source):
                return OCRResult("x" * 4097, source=source, confidence=0.5)
        with self.assertRaises(ValueError):
            ocr_docx_image(path, rel, OversizedOCR())
        self.assertEqual(path.read_bytes(), before)


if __name__ == "__main__":
    unittest.main()
