"""Regression tests for sparse tables, stale worksheet sizes and typed values."""

import sqlite3
import tempfile
import unittest
from pathlib import Path
from zipfile import ZipFile

from openpyxl import Workbook
from openpyxl.worksheet.formula import ArrayFormula, DataTableFormula

from converter import columns_for, export_table, read_collections


class ImportFidelityTests(unittest.TestCase):
    def setUp(self):
        folder = tempfile.TemporaryDirectory()
        self.addCleanup(folder.cleanup)
        self.folder = Path(folder.name)

    def sql_result(self, source):
        records = next(iter(read_collections(source).values()))
        target = self.folder / "result.sql"
        export_table(records, columns_for(records), target)
        db = sqlite3.connect(":memory:")
        self.addCleanup(db.close)
        db.executescript(target.read_bytes().decode("utf-8"))
        self.assertEqual(db.execute("PRAGMA integrity_check").fetchone(), ("ok",))
        return ([field[1] for field in db.execute("PRAGMA table_info(data)")],
                db.execute("SELECT * FROM data ORDER BY rowid").fetchall())

    def test_csv_retains_headers_even_when_all_records_are_short(self):
        source = self.folder / "short.csv"
        source.write_text("sku,name,notes\n001\n002,Tea\n", encoding="utf-8")
        self.assertEqual(self.sql_result(source),
                         (["sku", "name", "notes"], [("001", None, None), ("002", "Tea", None)]))

    def test_worksheet_dimensions_do_not_hide_data_or_invent_columns(self):
        base = self.folder / "original.xlsx"
        rows = [(f"{i:06}", f"Café 日本語 {i}") for i in range(1, 5002)]
        book = Workbook()
        book.active.append(["sku", "name"])
        for row in rows:
            book.active.append(row)
        book.save(base)
        book.close()

        for bounds in (b"A1:A2", b"A1:H5500"):
            with self.subTest(bounds=bounds):
                source = self.folder / "stale.xlsx"
                with ZipFile(base) as original, ZipFile(source, "w") as rewritten:
                    for entry in original.infolist():
                        data = original.read(entry)
                        if entry.filename == "xl/worksheets/sheet1.xml":
                            dimension = b'<dimension ref="A1:B5002"'
                            self.assertIn(dimension, data)
                            data = data.replace(dimension, b'<dimension ref="' + bounds + b'"', 1)
                        rewritten.writestr(entry, data)
                self.assertEqual(self.sql_result(source), (["sku", "name"], rows))

    def test_excel_keeps_blank_rows_and_unnamed_physical_columns(self):
        source = self.folder / "ragged.xlsx"
        book = Workbook()
        sheet = book.active
        sheet.append(["sku"])
        sheet.append(["001"])
        sheet["C4"] = "far cell"
        book.save(source)
        book.close()
        self.assertEqual(self.sql_result(source), (["sku", "Column 2", "Column 3"],
                         [("001", None, None), (None, None, None), (None, None, "far cell")]))

    def test_excel_round_trip_distinguishes_empty_text_from_null(self):
        source = self.folder / "empty-text.xlsx"
        records = [{"": "", "label": None}, {"": None, "label": ""}]
        export_table(records, ["", "label"], source)
        self.assertEqual(self.sql_result(source), (["", "label"], [("", None), (None, "")]))

    def test_excel_array_formula_keeps_its_expression(self):
        source = self.folder / "formula.xlsx"
        book = Workbook()
        book.active.append(["result"])
        book.active["A2"] = ArrayFormula(ref="A2:A3", text="=ROW(A2:A3)")
        book.save(source)
        book.close()
        self.assertEqual(self.sql_result(source), (["result"], [("=ROW(A2:A3)",)]))

    def test_excel_data_table_without_expression_explains_how_to_convert(self):
        source = self.folder / "data-table.xlsx"
        book = Workbook()
        book.active.append(["result"])
        book.active["A2"] = DataTableFormula(ref="A2:A3")
        book.save(source)
        book.close()
        with self.assertRaisesRegex(ValueError, "A2.*values"):
            read_collections(source)

    def test_typed_xml_checks_field_names_and_json_values(self):
        source = self.folder / "typed.xml"
        for fields, message in (
            ('<field name="x">1</field><field name="x">2</field>', "duplicate"),
            ('<field name="x">{"n":1,"n":2}</field>', "duplicate"),
            ('<field name="x">1e-400</field>', "precision"),
            ('<field name="x">NaN</field>', "valid JSON number"),
            (r'<field name="x">"\ud800"</field>', "Unicode"),
        ):
            with self.subTest(fields=fields):
                source.write_text(f'<records format="json-table-converter"><record>{fields}</record></records>',
                                  encoding="utf-8")
                with self.assertRaisesRegex(ValueError, message):
                    read_collections(source)
