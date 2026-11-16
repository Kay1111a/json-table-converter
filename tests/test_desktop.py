from pathlib import Path
import tempfile
import tkinter as tk
import unittest
from unittest.mock import patch

from desktop import ConverterApp


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
        with patch("desktop.filedialog.askopenfilename", return_value=str(example)):
            app.open_file()

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
