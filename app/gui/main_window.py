"""
Main application window — layout and event wiring only.
Business logic lives in report_actions.py and config_persistence.py.
"""
import os
import tkinter as tk
from tkinter import ttk, messagebox

from app.gui.file_dialogs import select_excel_file, select_docx_template, select_mapping_file
from app.gui.line_selector import LineSelector
from app.gui.config_persistence import (
    restore_config, save_config, save_mapping_path, load_mapping_config,
)
from app.gui.report_actions import generate_report

DEFAULT_MAPPING = os.path.abspath(
    os.path.join(os.path.dirname(__file__), "..", "mappings", "data.json")
)


class MainWindow(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title("Report Generator")
        self.geometry("600x420")
        self.resizable(False, False)

        self.excel_path    = tk.StringVar()
        self.template_path = tk.StringVar()
        self.mapping_path  = tk.StringVar()
        self.line_number   = tk.IntVar(value=2)

        self._build_ui()
        self._restore()

    # ------------------------------------------------------------------
    # UI construction
    # ------------------------------------------------------------------

    def _build_ui(self):
        ttk.Label(self, text="Report Generator", font=("Arial", 16, "bold")).pack(pady=(14, 8))

        self._path_row("Data file (.xlsx)",  "Select…", self.on_select_excel,    self.excel_path)
        self._path_row("Template (.docx)",   "Select…", self.on_select_template, self.template_path)
        self._path_row("Mapping (.json)",    "Select…", self.on_select_mapping,  self.mapping_path)

        ttk.Separator(self, orient="horizontal").pack(fill="x", padx=14, pady=10)

        LineSelector(self, variable=self.line_number).pack(padx=14, pady=4)

        ttk.Separator(self, orient="horizontal").pack(fill="x", padx=14, pady=10)

        ttk.Button(
            self, text="Generate report",
            command=self.on_generate_report,
        ).pack(pady=(4, 16), ipadx=20, ipady=6)

    def _path_row(self, label: str, btn_text: str, cmd, var: tk.StringVar):
        """One row: label | button | path display."""
        frame = ttk.Frame(self)
        frame.pack(fill="x", padx=14, pady=3)
        ttk.Label(frame, text=label, width=18, anchor="w").pack(side="left")
        ttk.Button(frame, text=btn_text, command=cmd, width=10).pack(side="left", padx=(0, 8))
        ttk.Label(frame, textvariable=var, foreground="#1a5fb4",
                  wraplength=320, anchor="w").pack(side="left", fill="x", expand=True)

    # ------------------------------------------------------------------
    # Startup restore
    # ------------------------------------------------------------------

    def _restore(self):
        state = restore_config()
        self.excel_path.set(state["excel"])
        self.template_path.set(state["docx"])
        self.line_number.set(state["line"])

        mapping = state["mapping"] or DEFAULT_MAPPING
        self.mapping_path.set(mapping)
        self._apply_mapping_config(mapping)

    # ------------------------------------------------------------------
    # Button callbacks
    # ------------------------------------------------------------------

    def on_select_excel(self):
        path = select_excel_file()
        if path:
            self.excel_path.set(path)

    def on_select_template(self):
        path = select_docx_template()
        if path:
            self.template_path.set(path)

    def on_select_mapping(self):
        path = select_mapping_file()
        if path:
            self.mapping_path.set(path)
            save_mapping_path(path)
            self._apply_mapping_config(path)

    def _apply_mapping_config(self, mapping_path: str):
        """Auto-fill Excel and template fields from the mapping's config block."""
        cfg = load_mapping_config(mapping_path)
        if cfg.get("data_file"):
            self.excel_path.set(cfg["data_file"])
        if cfg.get("template_file"):
            self.template_path.set(cfg["template_file"])

    def on_generate_report(self):
        excel   = self.excel_path.get()
        docx    = self.template_path.get()
        mapping = self.mapping_path.get()
        line    = self.line_number.get()

        if not excel or not docx:
            messagebox.showerror("Missing files", "Select an Excel file and a template first.")
            return
        if not mapping:
            messagebox.showerror("Missing mapping", "Select a mapping file first.")
            return
        if line < 2:
            messagebox.showerror("Invalid line", "Row number must be ≥ 2 (row 1 is the header).")
            return

        try:
            output_path = generate_report(excel, docx, mapping, row_number=line)
        except Exception as exc:
            messagebox.showerror("Generation error", str(exc))
            return

        save_config(excel, docx, mapping, line)
        self.line_number.set(line + 1)

        messagebox.showinfo("Done", f"Report generated:\n{output_path}")


if __name__ == "__main__":
    app = MainWindow()
    app.mainloop()
