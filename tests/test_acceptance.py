"""Windows acceptance tests with deterministic synthetic inputs and real exports."""

import json
from pathlib import Path
import sqlite3
from tempfile import TemporaryDirectory
import tkinter as tk
from tkinter import font as tkfont
import unittest
from unittest.mock import patch

from openpyxl import Workbook, load_workbook

from converter import INPUT_FORMATS, OUTPUT_FORMATS, columns_for, export_table, read_collections
from desktop import ConverterApp


HEADERS = ("sku", "product", "quantity", "price", "active", "notes")
ROW_COUNT = 5000


def make_rows():
    return [
        (f"{n:06}", f"Produit {n:05} · Café 日本語 ñ", n % 97,
         n / 8 + 0.125, n % 2 == 0, "O'Brien\nline two" if n % 3 == 0 else None)
        for n in range(1, ROW_COUNT + 1)
    ]


class AcceptanceTests(unittest.TestCase):
    def setUp(self):
        folder = TemporaryDirectory()
        self.addCleanup(folder.cleanup)
        self.path = Path(folder.name)

    def make_app(self):
        try:
            root = tk.Tk()
        except tk.TclError as exc:
            if "display" in str(exc).lower():
                self.skipTest("No graphical display is available.")
            raise
        self.addCleanup(root.destroy)
        root.withdraw()
        return ConverterApp(root)

    def write_json(self, filename, payload):
        path = self.path / filename
        path.write_text(json.dumps(payload, ensure_ascii=True), encoding="utf-8")
        return path

    @staticmethod
    def open_in_app(app, path):
        with patch("desktop.filedialog.askopenfilename", return_value=str(path)):
            app.open_file()

    def test_5000_json_and_excel_upload_preview_and_sql(self):
        app = self.make_app()
        rows = make_rows()
        source_json = self.write_json("inventory.json", [dict(zip(HEADERS, row)) for row in rows])
        source_xlsx = self.path / "inventory.xlsx"
        workbook = Workbook(write_only=True)
        sheet = workbook.create_sheet("Inventory")
        sheet.append(HEADERS)
        for row in rows:
            sheet.append(row)
        workbook.save(source_xlsx)
        workbook.close()

        for source in (source_json, source_xlsx):
            with self.subTest(source=source.name):
                with patch("desktop.messagebox.showerror") as load_error:
                    self.open_in_app(app, source)
                    load_error.assert_not_called()
                self.assertEqual(len(app.records), ROW_COUNT)
                self.assertEqual(tuple(app.columns), HEADERS)
                headings = [app.table.heading(i, "text") for i in app.table["columns"]]
                self.assertEqual(tuple(headings), HEADERS)
                app._select_columns(False)
                self.assertTrue(app.export_button.instate(["disabled"]))
                app._select_columns(True)
                self.assertTrue(app.export_button.instate(["!disabled"]))

                app.output_format.set("SQLite SQL (.sql)")
                sql_file = self.path / (source.suffix[1:] + "-inventory.sql")
                with patch("desktop.filedialog.asksaveasfilename", return_value=str(sql_file)), patch(
                    "desktop.messagebox.showerror"
                ) as export_error:
                    app.export_button.invoke()
                    export_error.assert_not_called()
                script = sql_file.read_bytes().decode("utf-8")
                self.assertEqual(script.count('INSERT INTO "data" VALUES'), ROW_COUNT)
                with sqlite3.connect(":memory:") as db:
                    db.executescript(script)
                    count = db.execute('SELECT COUNT(*) FROM "data"').fetchone()[0]
                    info = db.execute('PRAGMA table_info("data")').fetchall()
                    schema = [(col[1], col[2]) for col in info]
                    storage_types = db.execute(
                        'SELECT typeof(sku), typeof(quantity), typeof(price), '
                        'typeof(active), typeof(notes) FROM "data" WHERE rowid = 3'
                    ).fetchone()
                    integrity = db.execute('PRAGMA integrity_check').fetchone()[0]
                    self.assertEqual(count, ROW_COUNT)
                    self.assertEqual(integrity, "ok")
                    self.assertEqual(storage_types, ("text", "integer", "real", "integer", "text"))
                    for number, (actual, expected) in enumerate(zip(
                        db.execute('SELECT * FROM "data" ORDER BY rowid'), rows, strict=True
                    ), start=1):
                        self.assertEqual(actual, expected, f"{source.suffix} row {number}")
                    print(f"SQL proof {source.suffix}: INSERTs={count}, integrity={integrity}, schema={schema}", flush=True)
                    self.assertEqual(schema, list(zip(HEADERS, ("TEXT", "TEXT", "INTEGER", "REAL", "INTEGER", "TEXT"))))

    def test_preview_reaches_past_twenty_and_last_page(self):
        app = self.make_app()
        source = self.write_json("preview.json", [dict(zip(HEADERS, row)) for row in make_rows()])
        self.open_in_app(app, source)
        table = app.table
        initial = table.get_children()
        self.assertGreater(len(initial), 20, "Preview stops at its first 20 records.")
        self.assertIn("Produit 00021", table.item(initial[20], "values")[1])
        for _ in range(ROW_COUNT // 20):
            if app.next_page_button.instate(["disabled"]):
                break
            app.next_page_button.invoke()
        self.assertTrue(app.next_page_button.instate(["disabled"]))
        self.assertIn("Produit 05000", table.item(table.get_children()[-1], "values")[1])
        app.previous_page_button.invoke()
        self.assertNotIn("Produit 05000", table.item(table.get_children()[-1], "values")[1])

    def test_preview_reserves_width_for_long_headers(self):
        app = self.make_app()
        header = "Warehouse dispatch confirmation identifier"
        self.open_in_app(app, self.write_json("wide-header.json", [{header: "BX-101"}]))
        width = tkfont.Font(root=app.root, font=app.heading_font).measure(header)
        self.assertGreaterEqual(app.table.column("field0", "width"), width + 12)

    def test_bad_inputs_are_rejected_without_losing_current_preview(self):
        app = self.make_app()
        source = self.write_json("valid.json", [{"sku": "001"}])
        self.open_in_app(app, source)
        cases = {
            "empty.json": b"", "empty.xlsx": b"", "empty.csv": b"", "empty.xml": b"",
            "empty.bson": b"", "broken.json": b"{", "broken.xlsx": b"not a workbook",
            "broken.csv": b'header\n"unterminated', "broken.xml": b"<records>",
            "broken.bson": b"abc", "wrong.pdf": b"%PDF-1.4", "backup.dump": b"PGDMP",
        }
        for name, data in cases.items():
            with self.subTest(name=name):
                bad = self.path / name
                bad.write_bytes(data)
                with patch("desktop.messagebox.showerror") as error:
                    self.open_in_app(app, bad)
                    error.assert_called_once()
                self.assertEqual(app.source, source)
                self.assertEqual(app.records, [{"sku": "001"}])
                self.assertEqual(len(app.table.get_children()), 1)

    def test_invalid_json_unicode_is_rejected_before_replacing_preview(self):
        app = self.make_app()
        valid = self.write_json("valid.json", [{"sku": "001"}])
        invalid = self.write_json("invalid.json", [{"sku": "\ud800"}])
        self.open_in_app(app, valid)
        with patch("desktop.messagebox.showerror") as error:
            self.open_in_app(app, invalid)
            error.assert_called_once()
        self.assertEqual(app.source, valid)

    def test_duplicate_json_keys_are_rejected(self):
        source = self.path / "duplicate.json"
        source.write_text('[{"sku":"A","sku":"B"}]', encoding="utf-8")
        with self.assertRaisesRegex(ValueError, "duplicate"):
            read_collections(source)

    def test_json_numeric_loss_is_rejected_before_conversion(self):
        source = self.path / "precision.json"
        for literal in ("1e-400", "1.234567890123456789"):
            with self.subTest(literal=literal):
                source.write_text('[{"number":' + literal + '}]', encoding="utf-8")
                with self.assertRaisesRegex(ValueError, "precision"):
                    read_collections(source)

    def test_sql_wide_table_fails_before_replacing_existing_output(self):
        destination = self.path / "wide.sql"
        destination.write_text("previous output", encoding="utf-8")
        columns = [f"field_{i}" for i in range(2001)]
        with self.assertRaisesRegex(ValueError, "2,000"):
            export_table([{columns[0]: 1}], columns, destination)
        self.assertEqual(destination.read_text(encoding="utf-8"), "previous output")

    def test_sql_accepts_distinct_unicode_headers_and_preserves_mixed_types(self):
        destination = self.path / "names.sql"
        rows = [{"Ä": 1, "ä": 2.5, "mixed": "001"}, {"Ä": 2, "ä": 1.5, "mixed": 5}]
        export_table(rows, columns_for(rows), destination)
        with sqlite3.connect(":memory:") as db:
            db.executescript(destination.read_bytes().decode("utf-8"))
            self.assertEqual(db.execute('SELECT * FROM "data" ORDER BY rowid').fetchall(),
                             [(1, 2.5, "001"), (2, 1.5, 5)])

    def test_xlsx_does_not_round_values_excel_cannot_store_exactly(self):
        destination = self.path / "precision.xlsx"
        value = 1.2345678901234567
        export_table([{"price": value}], ["price"], destination)
        with destination.open("rb") as stream:
            workbook = load_workbook(stream)
            self.addCleanup(workbook.close)
            self.assertEqual(workbook.active["A2"].value, repr(value))

    def test_invalid_xml_character_cannot_create_broken_excel_export(self):
        destination = self.path / "invalid.xlsx"
        with self.assertRaisesRegex(ValueError, "unsupported character"):
            export_table([{"text": "A\ufffeZ"}], ["text"], destination)
        self.assertFalse(destination.exists())

    def test_excel_keeps_fully_blank_records_between_populated_rows(self):
        rows = [{"x": 1}, {"x": None}, {"x": 3}, {"x": None}]
        destination = self.path / "gaps.xlsx"
        export_table(rows, ["x"], destination)
        self.assertEqual(read_collections(destination)["Data"], rows)

    def test_converter_can_read_its_own_long_csv_cell(self):
        destination = self.path / "long.csv"
        rows = [{"body": "ñ" * 150000}]
        export_table(rows, ["body"], destination)
        self.assertEqual(read_collections(destination)["CSV rows"], rows)

    def test_xml_repeated_records_keep_empty_rows(self):
        source = self.path / "records.xml"
        source.write_text('<items><item><sku>A</sku></item><item/><item><sku>B</sku></item></items>',
                          encoding="utf-8")
        self.assertEqual(read_collections(source)["XML records"], [{"sku": "A"}, {}, {"sku": "B"}])

    def test_format_matrix_smoke(self):
        first = [{"id": "001", "name": "Café"}, {"id": "002", "name": "東京"}]
        for input_suffix in INPUT_FORMATS:
            source = self.path / ("input" + input_suffix)
            export_table(first, ["id", "name"], source)
            collections = read_collections(source)
            records = next(iter(collections.values()))
            self.assertEqual(records, first)
            for output_suffix in OUTPUT_FORMATS:
                with self.subTest(input=input_suffix, output=output_suffix):
                    target = self.path / ("output" + output_suffix)
                    export_table(records, ["id"], target)
                    if output_suffix == ".sql":
                        with sqlite3.connect(":memory:") as db:
                            db.executescript(target.read_bytes().decode("utf-8"))
                            self.assertEqual(db.execute('SELECT id FROM "data"').fetchall(), [("001",), ("002",)])
                    else:
                        self.assertEqual(next(iter(read_collections(target).values())),
                                         [{"id": "001"}, {"id": "002"}])
        print("Format matrix: 30 input/output pairs checked.", flush=True)

    def test_excel_round_trip_preserves_carriage_returns(self):
        rows = [{"notes": "first\r\nsecond\rlast\nfin"}]
        destination = self.path / "line-breaks.xlsx"
        export_table(rows, ["notes"], destination)
        self.assertEqual(read_collections(destination)["Data"], rows)
