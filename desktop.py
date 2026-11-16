"""Desktop interface for JSON Table Converter."""

from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinter import font as tkfont

from converter import columns_for, export_table, read_collections, table_value


APP_NAME = "JSON Table Converter"
PREVIEW_ROWS = 20

BACKGROUND = "#edf3f2"
PANEL = "#ffffff"
INK = "#102b2b"
MUTED = "#405957"
ACCENT = "#09665e"
BORDER = "#cddcda"


class ConverterApp:
    def __init__(self, root):
        self.root = root
        self.source = None
        self.collections = {}
        self.columns = []
        self.records = []

        root.title(APP_NAME)
        root.geometry("1020x650")
        root.minsize(760, 510)
        root.configure(background=BACKGROUND)
        root.columnconfigure(0, weight=1)
        root.rowconfigure(1, weight=1)

        family = "Inter" if "Inter" in tkfont.families() else tkfont.nametofont("TkDefaultFont").actual("family")
        self.body_font = (family, 11)
        self.heading_font = (family, 12, "bold")
        style = ttk.Style(root)
        style.theme_use("clam")
        style.configure("Action.TButton", font=self.heading_font, padding=(14, 8),
                        borderwidth=0, foreground=INK, background="#dceae8")
        style.map("Action.TButton", background=[("active", "#c6dfdb"), ("disabled", "#e6eceb")],
                  foreground=[("disabled", "#536966")])
        style.configure("Primary.TButton", font=self.heading_font, padding=(14, 9),
                        borderwidth=0, foreground=PANEL, background=ACCENT)
        style.map("Primary.TButton", background=[("active", "#07524e")])
        style.configure("TCombobox", padding=4)
        style.configure("Preview.Treeview", rowheight=29, font=self.body_font,
                        foreground=INK, background=PANEL, fieldbackground=PANEL)
        style.configure("Preview.Treeview.Heading", padding=(9, 7), font=self.heading_font,
                        foreground=INK, background="#e4eeec")
        style.map("Preview.Treeview.Heading", background=[("active", "#e4eeec")])
        style.map("Preview.Treeview", background=[("selected", "#d5e8e4")],
                  foreground=[("selected", INK)])

        header = tk.Frame(root, bg=BACKGROUND, padx=24, pady=19)
        header.grid(row=0, column=0, sticky="ew")
        header.columnconfigure(0, weight=1)
        tk.Label(header, text=APP_NAME, bg=BACKGROUND, fg=INK,
                 font=(family, 24, "bold")).grid(row=0, column=0, sticky="w")
        self.source_text = tk.StringVar(value="Open a JSON file to choose columns and preview the export.")
        tk.Label(header, textvariable=self.source_text, bg=BACKGROUND, fg=MUTED,
                 font=self.body_font, anchor="w").grid(row=1, column=0, sticky="ew")
        ttk.Button(header, text="Open JSON…", style="Primary.TButton",
                   command=self.open_file).grid(row=0, column=1, rowspan=2, padx=(18, 0))

        workspace = tk.Frame(root, bg=BACKGROUND)
        workspace.grid(row=1, column=0, sticky="nsew", padx=24)
        workspace.columnconfigure(1, weight=1)
        workspace.rowconfigure(0, weight=1)

        options = self._panel(workspace)
        options.grid(row=0, column=0, sticky="ns", padx=(0, 14))
        options.columnconfigure(0, weight=1)
        options.rowconfigure(5, weight=1)
        self._label(options, "Records", heading=True).grid(row=0, column=0, sticky="w")
        self.record_set = ttk.Combobox(options, state="disabled", width=22, font=self.body_font)
        self.record_set.grid(row=1, column=0, sticky="ew", pady=(5, 14))
        self.record_set.bind("<<ComboboxSelected>>", self._load_records)
        self._label(options, "Columns", heading=True).grid(row=2, column=0, sticky="w")
        self._label(options, "Click a field to include or exclude it.").grid(row=3, column=0, sticky="w")
        buttons = tk.Frame(options, bg=PANEL)
        buttons.grid(row=4, column=0, sticky="ew", pady=8)
        ttk.Button(buttons, text="Select all", style="Action.TButton",
                   command=lambda: self._select_columns(True)).pack(side="left")
        ttk.Button(buttons, text="Clear", style="Action.TButton",
                   command=lambda: self._select_columns(False)).pack(side="left", padx=(5, 0))

        list_frame = tk.Frame(options, bg=PANEL)
        list_frame.grid(row=5, column=0, sticky="nsew")
        field_scroll = ttk.Scrollbar(list_frame)
        self.fields = tk.Listbox(list_frame, selectmode="multiple", exportselection=False,
                                 yscrollcommand=field_scroll.set, font=self.body_font,
                                 fg=INK, bg="#f5f9f8", selectforeground=INK,
                                 selectbackground="#c7e3de", highlightcolor=ACCENT,
                                 highlightbackground=BORDER, relief="flat")
        field_scroll.configure(command=self.fields.yview)
        field_scroll.pack(side="right", fill="y")
        self.fields.pack(side="left", fill="both", expand=True)
        self.fields.bind("<<ListboxSelect>>", self._refresh_preview)

        preview = self._panel(workspace)
        preview.grid(row=0, column=1, sticky="nsew")
        preview.columnconfigure(0, weight=1)
        preview.rowconfigure(2, weight=1)
        self._label(preview, "Preview", heading=True).grid(row=0, column=0, sticky="w")
        self.preview_text = tk.StringVar(value="No file loaded")
        tk.Label(preview, textvariable=self.preview_text, font=self.body_font,
                 bg=PANEL, fg=MUTED).grid(row=1, column=0, sticky="w", pady=(1, 9))

        x_scroll = ttk.Scrollbar(preview, orient="horizontal")
        y_scroll = ttk.Scrollbar(preview)
        self.table = ttk.Treeview(preview, style="Preview.Treeview", show="headings",
                                  xscrollcommand=x_scroll.set, yscrollcommand=y_scroll.set)
        self.table.grid(row=2, column=0, sticky="nsew")
        x_scroll.configure(command=self.table.xview)
        x_scroll.grid(row=3, column=0, sticky="ew")
        y_scroll.configure(command=self.table.yview)
        y_scroll.grid(row=2, column=1, sticky="ns")

        footer = tk.Frame(root, bg=BACKGROUND, padx=24, pady=14)
        footer.grid(row=2, column=0, sticky="ew")
        footer.columnconfigure(0, weight=1)
        self.status = tk.StringVar()
        tk.Label(footer, textvariable=self.status, font=self.body_font,
                 bg=BACKGROUND, fg=MUTED, anchor="w", wraplength=420,
                 justify="left").grid(row=0, column=0, sticky="ew")
        self.export_buttons = []
        for index, (label, suffix) in enumerate((("Export CSV…", ".csv"), ("Export Excel…", ".xlsx")), 1):
            button = ttk.Button(footer, text=label, style="Action.TButton", state="disabled",
                                command=lambda extension=suffix: self.save_file(extension))
            button.grid(row=0, column=index, padx=(8, 0))
            self.export_buttons.append(button)

        root.bind("<Control-o>", lambda _event: self.open_file())

    def _panel(self, parent):
        return tk.Frame(parent, bg=PANEL, padx=15, pady=13,
                        highlightbackground=BORDER, highlightthickness=1)

    def _label(self, parent, text, *, heading=False):
        return tk.Label(parent, text=text, bg=PANEL, fg=INK if heading else MUTED,
                        font=self.heading_font if heading else self.body_font, anchor="w")

    def open_file(self):
        source = filedialog.askopenfilename(
            parent=self.root, title="Open JSON", filetypes=[("JSON", "*.json"), ("All files", "*")]
        )
        if not source:
            return
        try:
            collections = read_collections(source)
        except (OSError, ValueError, RecursionError) as exc:
            messagebox.showerror("Cannot open JSON", str(exc), parent=self.root)
            return

        self.source = Path(source)
        self.collections = collections
        self.source_text.set(f"File: {self.source.name}")
        self.record_set.configure(values=list(collections), state="readonly")
        self.record_set.current(1 if len(collections) == 2 and "Whole object" in collections else 0)
        self._load_records()

    def _load_records(self, _event=None):
        self.records = self.collections[self.record_set.get()]
        self.columns = columns_for(self.records)
        self.fields.delete(0, "end")
        for column in self.columns:
            self.fields.insert("end", column or "(empty field name)")
        self._select_columns(True)

    def _select_columns(self, selected):
        if selected:
            self.fields.selection_set(0, "end")
        else:
            self.fields.selection_clear(0, "end")
        self._refresh_preview()

    def _selected_columns(self):
        return [self.columns[index] for index in self.fields.curselection()]

    def _refresh_preview(self, _event=None):
        columns = self._selected_columns()
        self.table.delete(*self.table.get_children())
        identifiers = [f"field{index}" for index in range(len(columns))]
        self.table.configure(columns=identifiers)
        for identifier, name in zip(identifiers, columns):
            self.table.heading(identifier, text=name or "(empty field name)", anchor="w")
            self.table.column(identifier, anchor="w", width=155, minwidth=90)
        if columns:
            for index, row in enumerate(self.records[:PREVIEW_ROWS]):
                values = [table_value(row.get(column)) for column in columns]
                display = [" ".join(str(value).splitlines())[:120] if value is not None else ""
                           for value in values]
                self.table.insert("", "end", values=display, tags=("shaded",) if index % 2 else ())
            self.table.tag_configure("shaded", background="#f2f7f6")

        row_count = len(self.records)
        if not self.source:
            self.preview_text.set("No file loaded")
        elif not self.columns:
            self.preview_text.set("This record set contains no fields.")
        elif not columns:
            self.preview_text.set("Select columns to preview the export.")
        else:
            self.preview_text.set(f"Showing {min(PREVIEW_ROWS, row_count):,} of {row_count:,} rows")
        self.status.set(f"{len(columns):,} of {len(self.columns):,} columns selected" if self.source else "")
        for button in self.export_buttons:
            button.configure(state="normal" if columns and row_count else "disabled")

    def save_file(self, suffix):
        columns = self._selected_columns()
        if not self.source or not columns:
            return
        label = "CSV" if suffix == ".csv" else "Excel workbook"
        destination = filedialog.asksaveasfilename(
            parent=self.root, title=f"Export {label}", filetypes=[(label, f"*{suffix}")],
            defaultextension=suffix, initialfile=self.source.stem + suffix,
            initialdir=self.source.parent,
        )
        if not destination:
            return
        if Path(destination).suffix.lower() != suffix:
            messagebox.showerror("Invalid filename", f"Use the {suffix} extension.", parent=self.root)
            return
        try:
            self.root.configure(cursor="watch")
            self.root.update_idletasks()
            export_table(self.records, columns, destination)
        except (OSError, ValueError, RecursionError) as exc:
            messagebox.showerror("Export failed", str(exc), parent=self.root)
        else:
            self.status.set(f"Saved {len(self.records):,} rows to {Path(destination).name}")
        finally:
            self.root.configure(cursor="")


if __name__ == "__main__":
    window = tk.Tk()
    ConverterApp(window)
    window.mainloop()
