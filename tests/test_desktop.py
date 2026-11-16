from pathlib import Path
import tempfile
import tkinter as tk
import unittest
from unittest.mock import patch

from desktop import ConverterApp, upload_directory


class UploadDirectoryTests(unittest.TestCase):
    def test_prefers_downloads_then_documents_then_home(self):
        with tempfile.TemporaryDirectory() as folder:
            home = Path(folder)
            with patch("desktop.Path.home", return_value=home), patch(
                "desktop._windows_folder", return_value=None
            ):
                self.assertEqual(upload_directory(), home)
                documents = home / "Documents"
                documents.mkdir()
                self.assertEqual(upload_directory(), documents)
                downloads = home / "Downloads"
                downloads.mkdir()
                self.assertEqual(upload_directory(), downloads)

    def test_prefers_redirected_downloads_and_keeps_last_folder(self):
        with tempfile.TemporaryDirectory() as folder:
            home = Path(folder)
            redirected = home / "Redirected downloads"
            redirected.mkdir()
            last_folder = home / "Last upload"
            last_folder.mkdir()
            with patch("desktop.Path.home", return_value=home), patch(
                "desktop._windows_folder", return_value=redirected
            ):
                self.assertEqual(upload_directory(), redirected)
                self.assertEqual(upload_directory(last_folder / "input.json"), last_folder)


class DesktopTests(unittest.TestCase):
    def test_open_preview_clear_select_and_export(self):
        try:
            window = tk.Tk()
        except tk.TclError as exc:
            if "display" in str(exc).lower():
                self.skipTest("No graphical display is available.")
            raise
        self.addCleanup(window.destroy)
        window.withdraw()
        app = ConverterApp(window)
        self.assertTrue(app.export_buttons[0].instate(["disabled"]))

        example = Path(__file__).resolve().parents[1] / "examples" / "supplies.json"
        with patch("desktop.filedialog.askopenfilename", return_value=str(example)) as picker:
            app.open_file()
            self.assertEqual(picker.call_args.kwargs["initialdir"], str(upload_directory()))

        with patch("desktop.filedialog.askopenfilename", return_value="") as picker:
            app.open_file()
            self.assertEqual(picker.call_args.kwargs["initialdir"], str(example.parent))

        self.assertEqual(app.record_set.get(), "Array · supplies")
        self.assertEqual(len(app.table.get_children()), 4)
        self.assertTrue(app.export_buttons[0].instate(["!disabled"]))

        app._select_columns(False)
        self.assertTrue(app.export_buttons[0].instate(["disabled"]))
        app.fields.selection_set(0)
        app._refresh_preview()

        with tempfile.TemporaryDirectory() as folder:
            target = Path(folder) / "supplies.csv"
            with patch("desktop.filedialog.asksaveasfilename", return_value=str(target)):
                app.save_file(".csv")
            self.assertEqual(target.read_text(encoding="utf-8-sig").splitlines()[:2],
                             ["item_id", "00123"])
