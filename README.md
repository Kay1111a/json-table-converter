# JSON Table Converter

Convert a file locally on Windows. Preview the data, choose the columns you want, and save them in the format you select. The app does not send files to a server.

## Use the desktop app

Open the latest successful **Windows desktop app** run under **Actions**, download the `JsonTableConverter-Windows-x64` artifact, and extract it. Open `JsonTableConverter.exe` to run the app. The executable is portable and needs no separate Python installation.

1. Choose **Upload file**. The picker starts in Downloads, falls back to Documents or your home folder, and remembers your last folder until the app closes. Try the [sample supplies file](examples/supplies.json) first.
2. The app chooses an input reader from the file extension and checks that it can read the file. Pick the record set or worksheet if there is more than one. Select which columns to keep and check the preview.
3. Pick a format under **Convert to**, click **Convert**, and save the output.

## File formats

| Format | Import | Export |
| --- | --- | --- |
| JSON (`.json`) | A JSON object, an array of objects, or an array directly inside an object | A JSON array of records |
| CSV (`.csv`) | UTF-8 comma-separated data with headers | UTF-8 comma-separated data with headers |
| Excel (`.xlsx`) | Select one worksheet with headers and data | One worksheet named Data |
| XML (`.xml`) | A single record or repeated records with child fields or attributes | A record list with named fields and values encoded as JSON |
| BSON (`.bson`) | A stream of MongoDB BSON documents for one collection | A stream of BSON documents |
| SQLite SQL (`.sql`) | — | A script that creates and populates one SQLite table named `data` |

For SQL dumps and native database backups such as `.bak`, `.dump`, and `.backup`, restore the backup with the appropriate database tools, then export the table or collection you want in one of the supported input formats. SQLite SQL output contains the selected rows and columns. Original database schemas, indexes, and relationships require the database's own backup tools.

CSV values are read as text, and Excel exports preserve text IDs with leading zeros. Excel date and time values are imported as ISO-formatted text, and formula cells are imported as formula text. BSON types such as ObjectId and binary data remain native in BSON output and use Extended JSON in text formats. Missing values become blank cells, JSON nulls, or SQL NULL depending on the output.

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
