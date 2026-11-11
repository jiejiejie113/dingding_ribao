"""无需数据库的 Excel 导出验证：python -m unittest -v。"""
import tempfile
import unittest
from decimal import Decimal
from pathlib import Path
from unittest.mock import patch

from openpyxl import load_workbook

import export_mysql as exporter


class FakeCursor:
    description = [(name,) for name in ["id", "text", "decimal", "binary", "long"]]

    def __init__(self, rows):
        self.rows = list(rows)

    def __enter__(self):
        return self

    def __exit__(self, *args):
        pass

    def execute(self, sql):
        assert sql == "SELECT * FROM `a``b`"

    def fetchmany(self, size):
        batch, self.rows = self.rows[:size], self.rows[size:]
        return batch


class FakeConnection:
    def __init__(self, rows):
        self.rows = rows

    def cursor(self, *args):
        return FakeCursor(self.rows)


class ExportTests(unittest.TestCase):
    def test_roundtrip_and_split(self):
        rows = [
            (1234567890123456789, "=1+1\x00", Decimal("1.2300"), b"abc", "x" * 33000),
            (2, None, None, None, None),
            (3, "ok", None, None, None),
        ]
        with tempfile.TemporaryDirectory() as directory, patch.object(exporter, "MAX_DATA_ROWS", 2):
            path = Path(directory) / "result.xlsx"
            stats = exporter.export_table(FakeConnection(rows), "a`b", path, 2)
            self.assertEqual(stats, dict(rows=3, sheets=2, cleaned=1, truncated=1))
            workbook = load_workbook(path)
            try:
                self.assertEqual(len(workbook.worksheets), 2)
                sheet = workbook.worksheets[0]
                self.assertEqual(sheet.cell(2, 1).value, "1234567890123456789")
                self.assertEqual(sheet.cell(2, 2).value, "=1+1")
                self.assertEqual(sheet.cell(2, 2).data_type, "s")
                self.assertEqual(sheet.cell(2, 3).value, "1.2300")
                self.assertEqual(sheet.cell(2, 4).value, "base64:YWJj")
                self.assertEqual(len(sheet.cell(2, 5).value), 32767)
                self.assertEqual(workbook.worksheets[1].cell(2, 1).value, 3)
            finally:
                workbook.close()

    def test_empty_table_has_headers(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "empty.xlsx"
            stats = exporter.export_table(FakeConnection([]), "a`b", path, 2)
            self.assertEqual(stats["rows"], 0)
            workbook = load_workbook(path)
            try:
                self.assertEqual(workbook.active.max_row, 1)
                self.assertEqual(workbook.active.cell(1, 1).value, "id")
            finally:
                workbook.close()


if __name__ == "__main__":
    unittest.main()
