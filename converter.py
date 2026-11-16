"""Read JSON records and export selected fields as CSV or Excel."""

import csv
import json
import math
import os
from pathlib import Path
import tempfile


def read_collections(source):
    """Return selectable record sets from an array or object at the JSON root."""
    with Path(source).open(encoding="utf-8-sig") as stream:
        document = json.load(stream, parse_constant=_invalid_number, parse_float=_finite_float)
    if isinstance(document, list):
        return {"Top-level array": _check_records(document)}
    if not isinstance(document, dict):
        raise ValueError("Use a JSON object or an array of JSON objects.")

    collections = {"Whole object": [document]}
    for key, value in document.items():
        if isinstance(value, list) and value and all(isinstance(row, dict) for row in value):
            collections[f"Array · {key}"] = value
    return collections


def _invalid_number(value):
    raise ValueError(f"{value} is not a valid JSON number.")


def _finite_float(value):
    number = float(value)
    if not math.isfinite(number):
        raise ValueError(f"Number {value} is outside the supported range.")
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
    if isinstance(value, (dict, list)):
        return json.dumps(value, ensure_ascii=False, separators=(",", ":"))
    return value


def export_table(records, columns, destination):
    """Write the complete output before replacing the destination."""
    if not columns:
        raise ValueError("Select at least one column to export.")

    destination = Path(destination)
    suffix = destination.suffix.lower()
    writers = {".csv": _write_csv, ".xlsx": _write_xlsx}
    if suffix not in writers:
        raise ValueError("Choose a filename ending in .csv or .xlsx.")

    with tempfile.NamedTemporaryFile(
        dir=destination.parent, prefix=".json-table-", suffix=suffix, delete=False
    ) as stream:
        temporary = Path(stream.name)

    try:
        writers[suffix](records, columns, temporary)
        os.replace(temporary, destination)
    finally:
        temporary.unlink(missing_ok=True)


def _write_csv(records, columns, path):
    with path.open("w", encoding="utf-8-sig", newline="") as stream:
        writer = csv.writer(stream)
        writer.writerow(columns)
        writer.writerows([table_value(row.get(column)) for column in columns] for row in records)


def _write_xlsx(records, columns, path):
    from openpyxl import Workbook
    from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE

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
            if isinstance(value, str):
                if len(value) > 32_767 or ILLEGAL_CHARACTERS_RE.search(value):
                    raise ValueError(
                        f"Excel cannot store the text at row {row_number}, field {column!r}. "
                        "It contains an unsupported character or exceeds 32,767 characters."
                    )
                cell = sheet.cell(row_number, column_number, value)
                cell.data_type = "s"
            elif value is not None:
                sheet.cell(row_number, column_number, value)

    try:
        write_row(1, columns)
        for row_number, row in enumerate(records, start=2):
            write_row(row_number, (row.get(column) for column in columns))
        workbook.save(path)
    finally:
        workbook.close()
