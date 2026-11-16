"""Load tabular data and convert selected columns between file formats."""

import csv
from datetime import date, time, timedelta
from decimal import Decimal, InvalidOperation
import json
import math
import os
from pathlib import Path
import re
import tempfile
from threading import Lock
from xml.etree import ElementTree

from bson import decode_file_iter, encode, json_util
from bson.codec_options import DEFAULT_CODEC_OPTIONS, DatetimeConversion
from bson.errors import BSONError
from defusedxml import ElementTree as SafeXML
from defusedxml.common import DefusedXmlException
from openpyxl import Workbook, load_workbook
from openpyxl.utils.exceptions import InvalidFileException
from zipfile import BadZipFile


INPUT_FORMATS = {
    ".json": "JSON",
    ".csv": "CSV",
    ".xlsx": "Excel",
    ".xml": "XML",
    ".bson": "BSON",
}
OUTPUT_FORMATS = {**INPUT_FORMATS, ".sql": "SQLite SQL"}
BSON_OPTIONS = DEFAULT_CODEC_OPTIONS.with_options(datetime_conversion=DatetimeConversion.DATETIME_AUTO)
XML_FORMAT = "json-table-converter"
CSV_READ_LOCK = Lock()
SQLITE_COLUMN_LIMIT = 2_000
SQLITE_IDENTIFIER_FOLD = str.maketrans("ABCDEFGHIJKLMNOPQRSTUVWXYZ", "abcdefghijklmnopqrstuvwxyz")
INVALID_XML = re.compile(r"[^\x09\x0a\x0d\x20-\ud7ff\ue000-\ufffd\U00010000-\U0010ffff]")


def read_collections(source):
    """Choose a reader by file extension and validate the file as that format."""
    path = Path(source)
    suffix = path.suffix.lower()
    readers = {".json": _read_json, ".csv": _read_csv, ".xlsx": _read_xlsx,
               ".xml": _read_xml, ".bson": _read_bson}
    if suffix in {".sql", ".bak", ".dump", ".backup"}:
        raise ValueError("For database dumps and backups, export a table as CSV, JSON, XLSX, XML, or BSON first.")
    if suffix not in readers:
        raise ValueError("Choose a JSON, CSV, XLSX, XML, or BSON file.")

    try:
        return readers[suffix](path)
    except (BSONError, BadZipFile, csv.Error, DefusedXmlException,
            ElementTree.ParseError, InvalidFileException, KeyError, IndexError, OverflowError) as exc:
        raise ValueError(f"Cannot read this {INPUT_FORMATS[suffix]} file: {exc}") from exc


def _read_json(path):
    with path.open(encoding="utf-8-sig") as stream:
        document = json.load(
            stream, parse_constant=_invalid_number, parse_float=_finite_float,
            object_pairs_hook=_unique_object,
        )
    _validate_unicode(document)
    if isinstance(document, list):
        return {"Top-level array": _check_records(document)}
    if not isinstance(document, dict):
        raise ValueError("Use a JSON object or an array of JSON objects.")

    collections = {"Whole object": [document]}
    for key, value in document.items():
        if isinstance(value, list) and value and all(isinstance(row, dict) for row in value):
            collections[f"Array · {key}"] = value
    return collections


def _rows_with_headers(rows):
    header = next(rows, ())
    fields = [str(value) if value is not None else f"Column {index}"
              for index, value in enumerate(header, start=1)]
    if not fields:
        return []
    if len(set(fields)) != len(fields):
        raise ValueError("The file contains duplicate column names.")
    records = []
    for number, row in enumerate(rows, start=2):
        if not row:
            continue
        if len(row) > len(fields):
            raise ValueError(f"Row {number} contains more fields than the header.")
        values = [value.isoformat() if isinstance(value, (date, time))
                  else str(value) if isinstance(value, timedelta) else value for value in row]
        records.append(dict(zip(fields, values)))
    return records


def _read_csv(path):
    with path.open(encoding="utf-8-sig", newline="") as stream, CSV_READ_LOCK:
        limit = csv.field_size_limit()
        file_size = os.fstat(stream.fileno()).st_size
        csv.field_size_limit(max(limit, min(file_size, 2**31 - 1)))
        try:
            records = _rows_with_headers(csv.reader(stream, strict=True))
        finally:
            csv.field_size_limit(limit)
    if not records:
        raise ValueError("The CSV needs a header and at least one data row.")
    return {"CSV rows": records}


def _read_xlsx(path):
    with path.open("rb") as stream:
        workbook = load_workbook(stream, read_only=True, data_only=False)
        try:
            collections = {}
            for sheet in workbook.worksheets:
                records = _rows_with_headers(sheet.iter_rows(values_only=True))
                if records:
                    collections[sheet.title] = records
        finally:
            workbook.close()
    if not collections:
        raise ValueError("The workbook needs a worksheet with headers and data.")
    return collections


def _read_xml(path):
    root = SafeXML.parse(path).getroot()
    if root.tag == "records" and root.get("format") == XML_FORMAT:
        rows = [{field.attrib["name"]: json.loads(field.text or "")
                 for field in row.findall("field")} for row in root.findall("record")]
    else:
        children = list(root)
        has_fields = [bool(len(child) or child.attrib) for child in children]
        repeated_tags = len({child.tag for child in children}) == 1
        if any(has_fields) and (all(has_fields) or repeated_tags):
            rows = [_xml_fields(child) for child in children]
        else:
            rows = [_xml_fields(root)]
    if not any(rows):
        raise ValueError("The XML needs a row with attributes or child fields.")
    return {"XML records": rows}


def _xml_fields(element):
    result = {f"@{key}": value for key, value in element.attrib.items()}
    for child in element:
        key = child.tag
        value = ElementTree.tostring(child, encoding="unicode") if len(child) or child.attrib else child.text or ""
        if key in result:
            if not isinstance(result[key], list):
                result[key] = [result[key]]
            result[key].append(value)
        else:
            result[key] = value
    return result


def _read_bson(path):
    with path.open("rb") as stream:
        records = list(decode_file_iter(stream, codec_options=BSON_OPTIONS))
    if not records:
        raise ValueError("The BSON file contains no documents.")
    return {"BSON documents": records}


def _unique_object(pairs):
    result = {}
    for key, value in pairs:
        if key in result:
            raise ValueError(f"The JSON object contains a duplicate field name: {key!r}.")
        result[key] = value
    return result


def _validate_unicode(value):
    if isinstance(value, str):
        try:
            value.encode("utf-8")
        except UnicodeEncodeError as exc:
            raise ValueError("The JSON contains an invalid Unicode character.") from exc
    elif isinstance(value, list):
        for item in value:
            _validate_unicode(item)
    elif isinstance(value, dict):
        for key, item in value.items():
            _validate_unicode(key)
            _validate_unicode(item)


def _invalid_number(value):
    raise ValueError(f"{value} is not a valid JSON number.")


def _finite_float(value):
    number = float(value)
    if not math.isfinite(number):
        raise ValueError(f"Number {value} is outside the supported range.")
    try:
        exact = Decimal(value) == Decimal(str(number))
    except InvalidOperation:
        exact = False
    if not exact:
        raise ValueError(f"Number {value} would lose precision. Store it as a JSON string to keep its digits.")
    return number


def _check_records(records):
    if not records:
        raise ValueError("The JSON array is empty.")
    for index, row in enumerate(records, start=1):
        if not isinstance(row, dict):
            raise ValueError(f"Array item {index} must be a JSON object.")
    return records


def columns_for(records):
    """Keep field names in their first-seen order across all records."""
    return list(dict.fromkeys(key for row in records for key in row))


def table_value(value):
    if type(value) in (str, int, bool, type(None)):
        return value
    if type(value) is float and math.isfinite(value):
        return value
    if isinstance(value, int) and not isinstance(value, bool):
        return int(value)
    return json_util.dumps(value, ensure_ascii=False, separators=(",", ":"))


def _selected_rows(records, columns):
    for row in records:
        yield {column: row.get(column) for column in columns}


def export_table(records, columns, destination):
    """Write the complete output before replacing the destination."""
    if not columns:
        raise ValueError("Select at least one column to export.")

    destination = Path(destination)
    suffix = destination.suffix.lower()
    writers = {".csv": _write_csv, ".xlsx": _write_xlsx, ".json": _write_json,
               ".xml": _write_xml, ".bson": _write_bson, ".sql": _write_sql}
    if suffix not in writers:
        raise ValueError(f"Choose a filename ending in one of: {', '.join(OUTPUT_FORMATS)}.")

    with tempfile.NamedTemporaryFile(
        dir=destination.parent, prefix=".json-table-", suffix=suffix, delete=False
    ) as stream:
        temporary = Path(stream.name)

    try:
        writers[suffix](records, columns, temporary)
        os.replace(temporary, destination)
    except (BSONError, OverflowError) as exc:
        raise ValueError(f"Cannot export as {OUTPUT_FORMATS[suffix]}: {exc}") from exc
    finally:
        temporary.unlink(missing_ok=True)


def _write_csv(records, columns, path):
    with path.open("w", encoding="utf-8-sig", newline="") as stream:
        writer = csv.writer(stream)
        writer.writerow(columns)
        writer.writerows([table_value(row.get(column)) for column in columns] for row in records)


def _write_json(records, columns, path):
    data = json_util.dumps(list(_selected_rows(records, columns)), ensure_ascii=False, indent=2)
    path.write_text(data + "\n", encoding="utf-8")


def _write_xml(records, columns, path):
    if any(INVALID_XML.search(column) for column in columns):
        raise ValueError("A column name contains characters that XML cannot store.")
    root = ElementTree.Element("records", {"format": XML_FORMAT})
    for row in _selected_rows(records, columns):
        record = ElementTree.SubElement(root, "record")
        for key, value in row.items():
            field = ElementTree.SubElement(record, "field", name=key)
            field.text = json_util.dumps(value, ensure_ascii=True)
    ElementTree.indent(root)
    ElementTree.ElementTree(root).write(path, encoding="utf-8", xml_declaration=True)


def _write_bson(records, columns, path):
    with path.open("wb") as stream:
        for row in _selected_rows(records, columns):
            stream.write(encode(row, codec_options=BSON_OPTIONS))


def _sql_storage_value(value):
    value = table_value(value)
    if isinstance(value, bool):
        return int(value)
    if isinstance(value, int) and not -(2**63) <= value < 2**63:
        return str(value)
    return value


def _sql_literal(value):
    if value is None:
        return "NULL"
    if isinstance(value, (int, float)):
        return str(value)
    if "\0" in value:
        return f"CAST(X'{value.encode('utf-8').hex()}' AS TEXT)"
    return "'" + value.replace("'", "''") + "'"


def _sql_column_type(records, column):
    types = set()
    integer_requires_exact_storage = False
    for record in records:
        value = _sql_storage_value(record.get(column))
        if value is not None:
            types.add(type(value))
            if type(value) is int and float(value) != value:
                integer_requires_exact_storage = True

    if types == {int, float} and not integer_requires_exact_storage:
        return "REAL"
    storage_types = {int: "INTEGER", float: "REAL", str: "TEXT"}
    # BLOB affinity retains the original storage classes in mixed columns.
    return storage_types[types.pop()] if len(types) == 1 else "BLOB"


def _write_sql(records, columns, path):
    if len(columns) > SQLITE_COLUMN_LIMIT:
        raise ValueError("SQLite supports up to 2,000 columns. Deselect some columns or choose another format.")
    folded = {column.translate(SQLITE_IDENTIFIER_FOLD) for column in columns}
    if len(folded) != len(columns) or any("\0" in column for column in columns):
        raise ValueError("SQLite needs distinct column names without null characters.")

    declarations = []
    for column in columns:
        declared = _sql_column_type(records, column)
        escaped = column.replace('"', '""')
        declarations.append(f'"{escaped}" {declared}')

    with path.open("w", encoding="utf-8", newline="\n") as stream:
        stream.write(f'BEGIN TRANSACTION;\nCREATE TABLE "data" ({", ".join(declarations)});\n')
        for row in _selected_rows(records, columns):
            values = ", ".join(_sql_literal(_sql_storage_value(value)) for value in row.values())
            stream.write(f'INSERT INTO "data" VALUES ({values});\n')
        stream.write("COMMIT;\n")


def _write_xlsx(records, columns, path):
    if len(records) >= 1_048_576 or len(columns) > 16_384:
        raise ValueError("The table is too large for one Excel worksheet. Try CSV or fewer columns.")

    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Data"
    sheet.freeze_panes = "A2"

    def write_row(row_number, values):
        for column_number, (column, value) in enumerate(zip(columns, values), start=1):
            value = table_value(value)
            if type(value) is int and abs(value) >= 10**15:
                value = str(value)
            if isinstance(value, float) and (
                float(format(value, ".15g")) != value or abs(value) >= 1e308
                or 0 < abs(value) < 2.2250738585072014e-308
            ):
                value = repr(value)
            if isinstance(value, str):
                if len(value) > 32_767 or INVALID_XML.search(value):
                    raise ValueError(
                        f"Excel cannot store the text at row {row_number}, field {column!r}. "
                        "It contains an unsupported character or exceeds 32,767 characters."
                    )
                cell = sheet.cell(row_number, column_number, value)
                cell.data_type = "s"
            else:
                sheet.cell(row_number, column_number, value)

    try:
        write_row(1, columns)
        for row_number, row in enumerate(records, start=2):
            write_row(row_number, (row.get(column) for column in columns))
        workbook.save(path)
    finally:
        workbook.close()
