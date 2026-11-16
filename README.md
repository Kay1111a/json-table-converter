# JSON Table Converter

Select a JSON file, preview its records, and export the columns you need as CSV or an Excel workbook. This Python desktop project grew out of a JSON-to-CSV lab and uses fictional school supply records as its example.

## Use the desktop app

Open the latest successful **Windows desktop app** run under **Actions**, download the `JsonTableConverter-Windows-x64` artifact, and extract it. Open `JsonTableConverter.exe` to run the app. The executable is portable and needs no separate Python installation.

1. Choose **Upload JSON** and select your file. Try [the sample supplies file](examples/supplies.json) first.
2. Pick a record set if the file contains more than one. Select which columns to keep and check the preview.
3. Choose **Export CSV** or **Export Excel** and save your file.

## Supported input

- A top-level array of objects becomes one row per object.
- A top-level object can become a single row. A nonempty array of objects directly inside it can also be selected as a separate record set.
- Columns follow their first appearance across the records. Missing fields and JSON nulls export as blank cells. Nested objects and arrays export as JSON text.
- Excel exports retain JSON strings as text, including leading zeros. Long integer IDs are also written as text to preserve their digits. CSV has no cell types, so use the Excel export if you plan to double-click the file in Excel and preserve those IDs.

## Run from source

Use Python 3.12 with Tkinter.

```sh
python -m pip install -r requirements.txt
python desktop.py
```

Run the tests from the repo root. The desktop smoke test uses Tk and skips if there is no graphical display.

```sh
python -m unittest discover -s tests -v
```

## Build on Windows

```powershell
python -m pip install -r requirements-build.txt
python -m PyInstaller --noconfirm --onefile --windowed --name JsonTableConverter desktop.py
```

The executable is written to `dist/JsonTableConverter.exe`. The GitHub Actions workflow uses the same command on a Windows runner.
