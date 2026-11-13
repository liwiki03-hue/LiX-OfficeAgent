"""Regression coverage for inactive OOXML protection declarations."""
import json
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from openpyxl import Workbook

from office_agent.documents import classify_document, inspect_xlsx
from office_agent.writes import Editor, ExcelBarChart, ExcelSet, GuardError, LegacyWordReplace

_OLE_HEADER = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"


class _OpenBackend:
    """A backend whose probe always passes, so path guards are what the test exercises."""

    def probe(self, application):
        return {"application": application, "available": True, "stage": "ok", "reason": ""}

    def read_word_paragraphs(self, path, *, limit):
        return [line.decode("utf-8")
                for line in Path(path).read_bytes()[8:].split(b"\n")][:limit]

    def replace_word_paragraph(self, path, *, index, old, new, expected_count):
        raise AssertionError("no legacy edit may be reached in these tests")


class ProtectionClassificationTests(unittest.TestCase):
    def test_default_workbook_is_readable_and_editable(self):
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / "normal.xlsx"
            wb = Workbook()
            wb.active["A1"] = 2
            wb.save(path)
            classification = classify_document(path)
            self.assertTrue(classification.readable)
            self.assertFalse(classification.protected)
            self.assertEqual(inspect_xlsx(path)["sheets"][0]["title"], "Sheet")
            Editor(Path(folder)).plan(ExcelSet(kind="excel_set", path=str(path),
                                         sheet="Sheet", cell="A1", value=3))

    def test_active_workbook_or_sheet_protection_stays_blocked(self):
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / "protected.xlsx"
            wb = Workbook()
            wb.active["A1"] = 2
            wb.security.lockStructure = True
            wb.save(path)
            self.assertTrue(classify_document(path).protected)
            with self.assertRaises(GuardError):
                Editor(Path(folder)).plan(ExcelSet(kind="excel_set", path=str(path),
                                             sheet="Sheet", cell="A1", value=3))
            with self.assertRaises(GuardError):
                Editor(Path(folder)).plan(ExcelBarChart(kind="excel_bar_chart", path=str(path),
                                                        sheet="Sheet", values_range="A1:A2", title="Blocked"))

            wb.security.lockStructure = False
            wb.active.protection.sheet = True
            wb.save(path)
            self.assertTrue(classify_document(path).protected)


class LegacyChannelPathGuardTests(unittest.TestCase):
    """The COM channel must inherit the workspace guards, not bypass them."""

    def test_legacy_edit_outside_the_authorized_root_is_refused(self):
        with tempfile.TemporaryDirectory() as folder:
            base = Path(folder)
            root, outside = base / "authorized", base / "outside"
            root.mkdir()
            outside.mkdir()
            foreign = outside / "foreign.doc"
            foreign.write_bytes(_OLE_HEADER + b"price is 5")
            editor = Editor(root, legacy_backend=_OpenBackend())
            with self.assertRaisesRegex(GuardError, "outside authorized directory"):
                editor.plan(LegacyWordReplace(path=str(foreign), old="price", new="amount"))
            with self.assertRaisesRegex(GuardError, "absolute path is required"):
                editor.plan(LegacyWordReplace(path="foreign.doc", old="price", new="amount"))
            with self.assertRaisesRegex(GuardError, "relative/traversal path rejected"):
                editor.plan(LegacyWordReplace(path=str(root / ".." / "outside" / "x.doc"),
                                              old="price", new="amount"))
            self.assertFalse(editor.audit.exists())
            self.assertEqual(foreign.read_bytes(), _OLE_HEADER + b"price is 5")

    def test_legacy_edit_refuses_a_symlinked_source(self):
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            target = root / "real.doc"
            target.write_bytes(_OLE_HEADER + b"price is 5")
            link = root / "link.doc"
            try:
                link.symlink_to(target)
            except OSError as exc:                      # ordinary users cannot always symlink
                self.skipTest(f"symlink unavailable: {exc}")
            editor = Editor(root, legacy_backend=_OpenBackend())
            with self.assertRaisesRegex(GuardError, "ordinary file required"):
                editor.plan(LegacyWordReplace(path=str(link), old="price", new="amount"))
            self.assertEqual(target.read_bytes(), _OLE_HEADER + b"price is 5")
            self.assertFalse(editor.audit.exists())

    def test_legacy_plan_never_touches_the_source_before_confirmation(self):
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            source = root / "plan-only.doc"
            source.write_bytes(_OLE_HEADER + b"First\nprice is 5\nLast")
            before = source.read_bytes()
            editor = Editor(root, legacy_backend=_OpenBackend())
            plan = editor.plan(LegacyWordReplace(path=str(source), paragraph_index=1,
                                                 old="price", new="amount"))
            self.assertTrue(plan.confirmation_required)
            self.assertEqual(plan.before, "price is 5")
            self.assertEqual(plan.after, "amount is 5")
            self.assertEqual(source.read_bytes(), before)
            self.assertFalse(editor.audit.exists())
