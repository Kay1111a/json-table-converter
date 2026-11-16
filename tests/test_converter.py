import csv
import json
from pathlib import Path
import sqlite3
import tempfile
import unittest

from bson import Binary, ObjectId
from openpyxl import Workbook, load_workbook

from converter import columns_for, export_table, read_collections


class ConverterTests(unittest.TestCase):
    def setUp(self):
        workspace = tempfile.TemporaryDirectory()
        self.addCleanup(workspace.cleanup)
        self.directory = Path(workspace.name)

    def read(self, content):
        source = self.directory / "input.json"
        source.write_text(content, encoding="utf-8-sig")
        return read_collections(source)

    def test_collection_choices_and_sparse_fields(self):
        result = self.read(json.dumps({"items": [{"sku": "001"}, {"qty": 3}], "note": "Demo"}))
        self.assertEqual(list(result), ["Whole object", "Array · items"])
        self.assertEqual(columns_for(result["Array · items"]), ["sku", "qty"])

    def test_root_array(self):
        self.assertEqual(self.read('[{"name":"A"}]'), {"Top-level array": [{"name": "A"}]})

    def test_invalid_input(self):
        for content, message in (
            ("[]", "empty"),
            ('[{}, "text"]', "item 2"),
            ("42", "JSON object"),
            ('{"quantity":NaN}', "valid JSON number"),
            ('{"quantity":1e400}', "outside the supported range"),
            ('{"missing":}', "Expecting value"),
        ):
            with self.subTest(content=content):
                with self.assertRaisesRegex(ValueError, message):
                    self.read(content)

    def test_csv_quoting_unicode_and_column_selection(self):
        destination = self.directory / "output.CSV"
        records = [{"id": "001", "name": "Papel, A4\nBaguió", "extra": 3}, {"id": "002"}]
        export_table(records, ["name", "id"], destination)
        self.assertTrue(destination.read_bytes().startswith(b"\xef\xbb\xbf"))
        with destination.open(encoding="utf-8-sig", newline="") as stream:
            self.assertEqual(list(csv.reader(stream)), [["name", "id"], ["Papel, A4\nBaguió", "001"], ["", "002"]])

    def test_xlsx_preserves_text_boolean_number_and_nested_json(self):
        destination = self.directory / "output.xlsx"
        records = [{"=text": "=1+1", "id": "001", "count": 2, "active": False,
                    "nested": {"color": "teal"}, "serial": 12345678901234567, "missing": None}]
        export_table(records, columns_for(records), destination)
        workbook = load_workbook(destination)
        self.addCleanup(workbook.close)
        sheet = workbook["Data"]
        self.assertEqual(sheet.freeze_panes, "A2")
        self.assertEqual(sheet["A1"].data_type, "s")
        self.assertEqual(sheet["A2"].data_type, "s")
        self.assertEqual(tuple(cell.value for cell in sheet[2]),
                         ("=1+1", "001", 2, False, '{"color":"teal"}', "12345678901234567", None))
        self.assertEqual(sheet["D2"].data_type, "b")
        self.assertEqual(sheet["F2"].data_type, "s")

    def test_failed_export_leaves_existing_file_intact(self):
        destination = self.directory / "output.xlsx"
        destination.write_bytes(b"existing")
        with self.assertRaisesRegex(ValueError, "row 2"):
            export_table([{"notes": "A" * 32768}], ["notes"], destination)
        self.assertEqual(destination.read_bytes(), b"existing")
        self.assertEqual(list(self.directory.iterdir()), [destination])

    def test_rejects_no_columns_and_unsupported_extension(self):
        with self.assertRaisesRegex(ValueError, "Select at least one"):
            export_table([{}], [], self.directory / "out.xlsx")
        with self.assertRaisesRegex(ValueError, "filename ending"):
            export_table([{"x": 1}], ["x"], self.directory / "out.backup")

    def test_csv_import_preserves_strings_and_quoted_fields(self):
        source = self.directory / "source.CSV"
        source.write_text('sku,name,notes\n001,"A, B","first\nsecond"\n002,C\n', encoding="utf-8-sig")
        records = read_collections(source)["CSV rows"]
        self.assertEqual(records, [
            {"sku": "001", "name": "A, B", "notes": "first\nsecond"},
            {"sku": "002", "name": "C"},
        ])
        target = self.directory / "selected.json"
        export_table(records, ["sku", "notes"], target)
        self.assertEqual(json.loads(target.read_text(encoding="utf-8")),
                         [{"sku": "001", "notes": "first\nsecond"}, {"sku": "002", "notes": None}])

    def test_excel_import_exposes_worksheets_and_keeps_formulas_as_text(self):
        source = self.directory / "sheets.xlsx"
        workbook = Workbook()
        workbook.active.title = "Stock"
        workbook.active.append(["id", "total"])
        workbook.active.append(["001", "=1+1"])
        workbook.create_sheet("Orders").append(["id"])
        workbook["Orders"].append([12])
        workbook.save(source)
        workbook.close()
        self.assertEqual(read_collections(source), {
            "Stock": [{"id": "001", "total": "=1+1"}],
            "Orders": [{"id": 12}],
        })

    def test_xml_import_and_typed_export(self):
        source = self.directory / "source.xml"
        source.write_text('<items><item code="001"><name>A &amp; B</name></item></items>', encoding="utf-8")
        self.assertEqual(read_collections(source)["XML records"], [{"@code": "001", "name": "A & B"}])
        records = [{"Fancy field &": "first\r\nsecond", "qty": 2, "nested": ["A", None]}]
        target = self.directory / "typed.xml"
        export_table(records, columns_for(records), target)
        self.assertEqual(read_collections(target)["XML records"], records)

    def test_bson_import_export_preserves_objectid_and_binary(self):
        source = self.directory / "source.bson"
        records = [{"id": ObjectId("64b64c152a9e4eb49bc7ac01"),
                    "attachment": Binary(b"abc", 128)}, {"id": ObjectId(), "attachment": None}]
        export_table(records, ["id", "attachment"], source)
        self.assertEqual(read_collections(source)["BSON documents"], records)
        target = self.directory / "source.json"
        export_table(records, ["id", "attachment"], target)
        result = json.loads(target.read_text(encoding="utf-8"))
        self.assertEqual(result[0]["id"], {"$oid": "64b64c152a9e4eb49bc7ac01"})
        self.assertEqual(result[0]["attachment"]["$binary"]["subType"], "80")

    def test_sql_export_rebuilds_one_sqlite_table(self):
        target = self.directory / "table.sql"
        records = [{'quoted"name': "O'Brien", "qty": 2, "sku": "001", "note": "a\0b", "status": None}]
        export_table(records, columns_for(records), target)
        database = sqlite3.connect(":memory:")
        self.addCleanup(database.close)
        database.executescript(target.read_text(encoding="utf-8"))
        self.assertEqual(database.execute('SELECT * FROM "data"').fetchall(),
                         [("O'Brien", 2, "001", "a\0b", None)])

    def test_rejects_invalid_inputs_and_database_backups(self):
        for filename, data, message in (
            ("corrupt.bson", b"abc", "Cannot read this BSON"),
            ("corrupt.xlsx", b"abc", "Cannot read this Excel"),
            ("corrupt.xml", b"<broken>", "Cannot read this XML"),
            ("duplicate.csv", b"id,id\n1,2", "duplicate column"),
            ("backup.bak", b"", "export a table"),
            ("database.sql", b"", "export a table"),
        ):
            with self.subTest(filename=filename):
                source = self.directory / filename
                source.write_bytes(data)
                with self.assertRaisesRegex(ValueError, message):
                    read_collections(source)
