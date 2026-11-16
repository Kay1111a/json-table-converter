import csv
import json
from pathlib import Path
import tempfile
import unittest

from openpyxl import load_workbook

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
            export_table([{"x": 1}], ["x"], self.directory / "out.json")
