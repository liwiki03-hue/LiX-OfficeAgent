"""Bounded deterministic spreadsheet analysis; never recalculates formulas."""
from __future__ import annotations

from pathlib import Path
from math import isfinite
from openpyxl import load_workbook
from docx import Document


def read_word_table_region(path: Path, table_index: int, start_row: int, start_column: int,
                           rows: int, columns: int) -> dict:
    if min(table_index, start_row, start_column) < 0 or not (1 <= rows <= 20 and 1 <= columns <= 20):
        raise ValueError("invalid bounded Word table region")
    doc = Document(path)
    if table_index >= len(doc.tables):
        raise ValueError("table index out of bounds")
    table = doc.tables[table_index]
    selected = []
    for row_index in range(start_row, min(start_row + rows, len(table.rows))):
        row = table.rows[row_index]
        # python-docx reconstructs the table grid for each row.cells access.
        # Resolve it once per selected row, not once per returned cell.
        cells = row.cells
        selected.append({"row": row_index, "cells": [
            {"column": col_index, "text": cells[col_index].text[:4096]}
            for col_index in range(start_column, min(start_column + columns, len(cells)))]})
    return {"table_index": table_index, "total_rows": len(table.rows),
            "total_columns": len(table.columns), "rows": selected}


def filter_sort_rows(path: Path, sheet: str, start_row: int, end_row: int,
                     start_column: int, columns: int, key_column: int,
                     equals: str | int | float | bool | None = None,
                     descending: bool = False) -> dict:
    """Return a bounded sorted/filtered *view*, without changing worksheet order.

    Formula rows are omitted; a formula's expression/cache is not a sortable value.
    Filtering compares exact literal values and is opt-in when equals is not None.
    """
    if not (1 <= start_row <= end_row <= 1048576 and end_row - start_row < 100
            and 1 <= start_column <= 16384 and 1 <= columns <= 20
            and start_column + columns <= 16385 and start_column <= key_column < start_column + columns):
        raise ValueError("view limited to 100 rows and 20 columns with a key inside the region")
    if isinstance(equals, float) and not isfinite(equals):
        raise ValueError("filter must be finite")
    if isinstance(equals, str) and len(equals) > 500:
        raise ValueError("filter text is too long")
    workbook = load_workbook(path, data_only=False, read_only=True)
    try:
        if sheet not in workbook:
            raise ValueError("worksheet not found")
        rows = []
        skipped_formulas = 0
        for cells in workbook[sheet].iter_rows(min_row=start_row, max_row=end_row,
                min_col=start_column, max_col=start_column + columns - 1):
            if any(cell.data_type == "f" for cell in cells):
                skipped_formulas += 1
                continue
            key = cells[key_column - start_column].value
            if equals is not None and (type(key) is not type(equals) or key != equals):
                continue
            def display(value):
                if type(value) is float and not isfinite(value):
                    return str(value)  # Never emit nonstandard NaN/Infinity JSON values.
                return value if isinstance(value, (str, int, float, bool, type(None))) else str(value)
            rows.append({"row": cells[0].row, "key": display(key),
                         "values": [display(cell.value) for cell in cells]})
        # Mixed data types have no natural Excel ordering: group by type, then value.
        def sort_key(item):
            value = item["key"]
            if value is None:
                return (0, "")
            if type(value) is int or (type(value) is float and isfinite(value)):
                return (1, value)
            if type(value) is bool:
                return (2, int(value))
            return (3, str(value))
        rows.sort(key=sort_key, reverse=descending)
        return {"sheet": sheet, "source_rows": [start_row, end_row],
                "source_columns": [start_column, start_column + columns - 1],
                "key_column": key_column, "descending": descending,
                "formula_rows_excluded": skipped_formulas, "read_only_view": True, "rows": rows}
    finally:
        workbook.close()


def summarize_column(path: Path, sheet: str, column: int, start_row: int, end_row: int) -> dict:
    if not (1 <= column <= 16384 and 1 <= start_row <= end_row <= 1048576) or end_row - start_row >= 1000:
        raise ValueError("summary accepts 1-1000 rows and a valid Excel column")
    workbook = load_workbook(path, data_only=False, read_only=True)
    try:
        if sheet not in workbook:
            raise ValueError("worksheet not found")
        rows = workbook[sheet].iter_rows(min_row=start_row, max_row=end_row, min_col=column, max_col=column)
        numbers: list[int | float] = []
        formulas = empty = other = 0
        for row in rows:
            cell = row[0]
            if cell.data_type == "f":
                formulas += 1
            elif type(cell.value) in (int, float) and (type(cell.value) is int or isfinite(cell.value)):
                numbers.append(cell.value)  # bool/non-finite values are not numeric measurements.
            elif cell.value is None:
                empty += 1
            else:
                other += 1
        total = sum(numbers)
        if isinstance(total, float) and not isfinite(total):
            raise ValueError("numeric sum exceeds finite range")
        return {"sheet": sheet, "column": column, "range_rows": [start_row, end_row],
                "numeric_literals_only": True, "formula_cells_excluded": formulas,
                "empty": empty, "other": other, "count": len(numbers),
                "sum": total, "average": total/len(numbers) if numbers else None,
                "minimum": min(numbers) if numbers else None,
                "maximum": max(numbers) if numbers else None}
    finally:
        workbook.close()
