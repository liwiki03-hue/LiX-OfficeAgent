import tempfile
import unittest
from pathlib import Path
from openpyxl import Workbook
from docx import Document
from office_agent.analytics import summarize_column, read_word_table_region, filter_sort_rows


class AnalyticsTests(unittest.TestCase):
    def test_filter_sort_is_read_only_and_excludes_formula_rows(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "view.xlsx"
            book = Workbook(); sheet = book.active; sheet.title = "Data"
            sheet.append(["b", 3]); sheet.append(["a", 2]); sheet.append(["=A1", 1])
            book.save(path)
            original = path.read_bytes()
            view = filter_sort_rows(path, "Data", 1, 3, 1, 2, 1)
            self.assertEqual([row["row"] for row in view["rows"]], [2, 1])
            self.assertEqual(view["formula_rows_excluded"], 1)
            self.assertTrue(view["read_only_view"])
            self.assertEqual(filter_sort_rows(path, "Data", 1, 3, 1, 2, 1, equals="b")["rows"][0]["row"], 1)
            self.assertEqual(original, path.read_bytes())
            with self.assertRaises(ValueError):
                filter_sort_rows(path, "Data", 1, 101, 1, 2, 1)

    def test_word_table_region_is_bounded_and_located(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "table.docx"
            document = Document()
            table = document.add_table(rows=2, cols=2)
            table.cell(1, 1).text = "selected"
            document.save(path)
            result = read_word_table_region(path, 0, 1, 1, 1, 1)
            self.assertEqual(result["rows"], [{"row": 1, "cells": [{"column": 1, "text": "selected"}]}])
            with self.assertRaises(ValueError):
                read_word_table_region(path, 0, 0, 0, 21, 1)

    def test_numeric_summary_excludes_uncalculated_formulas_and_booleans(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "data.xlsx"
            wb = Workbook(); ws = wb.active; ws.title = "Data"
            for row, value in enumerate([2, 3, "=SUM(A1:A2)", True, "unknown"], 1):
                ws.cell(row, 1).value = value
            wb.save(path)
            result = summarize_column(path, "Data", 1, 1, 5)
            self.assertEqual(result["sum"], 5)
            self.assertEqual(result["average"], 2.5)
            self.assertEqual(result["count"], 2)
            self.assertEqual(result["formula_cells_excluded"], 1)
            self.assertEqual(result["other"], 2)
            with self.assertRaises(ValueError):
                summarize_column(path, "Data", 1, 1, 1001)
