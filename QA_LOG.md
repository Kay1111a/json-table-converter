# Conversion QA log

The audit used synthetic JSON and Excel files with 5,000 rows each. The fixtures include Unicode, apostrophes, embedded line breaks, leading-zero IDs, booleans, decimals, and nulls. The desktop tests use real Tk widgets on a Windows runner, with only the native file dialogs and error messages substituted so the tests can pick files and observe failures. The app supports reading input and creating or replacing output files. It has no feature for editing or deleting source records.

- [Original reproduction](https://github.com/Kay1111a/json-table-converter/actions/runs/43963656893)
- [Bounded baseline reproduction](https://github.com/Kay1111a/json-table-converter/actions/runs/43964301105)
- [Data-integrity fix retest](https://github.com/Kay1111a/json-table-converter/actions/runs/43964570588)
- [Schema follow-up retest](https://github.com/Kay1111a/json-table-converter/actions/runs/43964925309)
- [Passing acceptance run](https://github.com/Kay1111a/json-table-converter/actions/runs/43965214680)

The Windows acceptance suite verified fixes for all 14 defects below. Six were Critical, six High, and two Medium. None of the reproduced defects remain blocked.

## Defects

| ID | Severity | Reproduction | Expected | Baseline result | Affected paths |
|---|---|---|---|---|---|
| QA-01 | Critical | 1. Create JSON with two `sku` keys in one object. 2. Upload it. | Reject the ambiguous input. | The last value silently replaced the first. | `converter.py` |
| QA-02 | Critical | 1. Create JSON containing `1e-400` or `1.234567890123456789`. 2. Upload and read the record. | Preserve the value or report that it cannot be represented exactly. | Input was accepted after underflow or rounding. | `converter.py` |
| QA-03 | Critical | 1. Export a table with blank middle and trailing rows to Excel. 2. Upload the workbook. | Preserve the row count and positions. | Fully blank rows were dropped from the table. | `converter.py` |
| QA-04 | Critical | 1. Upload repeated XML item elements with one empty item. 2. Preview the records. | Show a row per item, including the empty one. | Items collapsed into one row containing an array. | `converter.py` |
| QA-05 | Critical | 1. Export a record containing `1.2345678901234567` to Excel. 2. Read the cell. | Preserve the original digits. | The value became `1.234567890123457`. | `converter.py` |
| QA-06 | Critical | 1. Write Excel rows containing embedded CRLF. 2. Upload and convert to SQL. 3. Query the text value. | Preserve the original line breaks. | The original Windows run found that CRLF became two LF characters. The fix also checks Excel round-trip with carriage returns in the packaged app. | `converter.py`, `requirements.txt`, `desktop.py` |
| QA-07 | High | 1. Load a valid file. 2. Upload JSON containing an unpaired Unicode surrogate. | Show an error and keep the previous preview. | The invalid file was accepted and replaced the current input. | `converter.py`, `desktop.py` |
| QA-08 | High | 1. Export text containing U+FFFE to Excel. 2. Check whether the writer validates it. | Reject the unsupported character before creating output. | The writer accepted the character even though worksheet XML cannot store it. | `converter.py` |
| QA-09 | High | 1. Convert the JSON or Excel inventory to SQLite SQL. 2. Import it and run `PRAGMA table_info("data")`. | Column names and declared types should reflect the sample data. | Every declared type was blank. | `converter.py` |
| QA-10 | High | 1. Select 2,001 columns. 2. Export to an existing SQL path. | Explain the SQLite column limit and preserve the existing output. | The converter wrote a script too wide for standard SQLite and replaced the previous file. | `converter.py` |
| QA-11 | High | 1. Upload the 5,000-row JSON file. 2. Try to reach row 21 or the last row. | Make the remaining records visible. | The table displayed only the first 20 rows. | `desktop.py` |
| QA-12 | High | 1. Upload a field named Warehouse dispatch confirmation identifier. 2. Inspect the preview column width. | Give the header enough width and allow horizontal scrolling. | The column was 155 pixels wide when the test required at least 346. | `desktop.py` |
| QA-13 | Medium | 1. Export a table with distinct headers `Ä` and `ä` to SQLite SQL. 2. Read the result. | Keep both valid SQLite identifiers. | Validation rejected them as duplicate identifiers. | `converter.py` |
| QA-14 | Medium | 1. Export a CSV with a 150,000-character cell. 2. Upload the exported file. | Read the entire cell. | The CSV reader rejected its own export at the default 131,072-character limit. | `converter.py` |

## Reproduce the verification

On Windows, run:

```powershell
python -m pip install -r requirements.txt
python -m unittest discover -s tests -v
```

The acceptance suite opens each input through the desktop upload handler, checks the headers and column controls, and invokes Convert on the real Tk button. It checks the number of INSERT statements, runs each generated script in a fresh SQLite database using `sqlite3.Connection.executescript`, and verifies `SELECT COUNT(*) FROM "data"`, `PRAGMA integrity_check`, `PRAGMA table_info("data")`, storage types, and every row. It separately covers navigation to row 5,001 and runs all 30 supported input/output format pairs. Empty, corrupt, and wrong-format inputs must show an error while preserving the previously loaded data.

For both inventory files, the generated SQL contained 5,000 INSERT statements and the imported table contained 5,000 rows. SQLite reported `ok` for both integrity checks, and all rows matched the respective inputs. The schema for both was `sku TEXT, product TEXT, quantity INTEGER, price REAL, active INTEGER, notes TEXT`.

## Import and preview regressions

| Area | How to reproduce | Previous result | Regression check |
|---|---|---|---|
| CSV headers | 1. Give a CSV three headers. 2. Omit the last field in every row. 3. Convert to SQL. | The last column disappeared. | SQLite contains all three columns and returns NULL for missing values. |
| Worksheet bounds | 1. Give a 5,001-row sheet a saved range of A1:A2 or A1:H5500. 2. Upload and convert it. | A short range truncated rows and columns. A wide range invented columns. | Both scripts import with the two original headers, all 5,001 rows in order, and an integrity check of `ok`. |
| Empty Excel text | 1. Export empty strings, a blank field name and nulls to XLSX. 2. Reopen and convert to SQL. | Empty text became null and the field name changed. | SQLite returns the original header and distinct empty strings and NULLs. |
| Excel formulas | 1. Upload an array formula or a what-if data table. 2. Inspect the export. | Control attributes appeared in place of the formula or its result. | An array formula retains its expression. A data table without formula text reports the cell and how to convert it. |
| Typed XML | 1. Upload converter XML with a duplicate field, an invalid Unicode escape or a number that loses precision. | Input was accepted with lost or invalid values. | Upload rejects the record with a numbered error. The same JSON validation applies to typed XML field values. |
| Preview widths | 1. Widen a column. 2. Go to the next page and back. | Each page reset the width. | The resized width and the correct page contents survive navigation. |
| Worksheet choice | 1. Create two sheets with the first named `Whole object`. 2. Upload the file. | The second sheet opened automatically. | The first sheet opens in workbook order. |

The data checks run without a display in `tests/test_import_fidelity.py`. The two preview checks run against real Tk widgets in `tests/test_acceptance.py`.
