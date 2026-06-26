import os
import datetime
import tkinter as tk
from tkinter import ttk, messagebox

from app.core.report_generator import ReportGenerator
from app.gui.file_dialogs import select_excel_file, select_docx_template, select_mapping_file
from app.repository.config_repository import config_get, config_set
from app.repository.rapport_repository import save_report

# Config keys
KEY_EXCEL   = "last_excel_path"
KEY_DOCX    = "last_docx_path"
KEY_MAPPING = "last_mapping_path"
KEY_LINE    = "last_line_number"

# Default mapping (fallback when nothing is stored yet)
DEFAULT_MAPPING = os.path.abspath(
    os.path.join(os.path.dirname(__file__), "..", "mappings", "data.json")
)


class MainWindow(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title("Report Generator")
        self.geometry("600x420")
        self.resizable(False, False)

        # Tkinter variables
        self.excel_path   = tk.StringVar()
        self.template_path = tk.StringVar()
        self.mapping_path  = tk.StringVar()
        self.line_number   = tk.IntVar(value=1)

        self._build_ui()
        self._restore_config()

    # ------------------------------------------------------------------
    # UI construction
    # ------------------------------------------------------------------

    def _build_ui(self):
        pad = {"padx": 14, "pady": 4}

        ttk.Label(self, text="Report Generator", font=("Arial", 16, "bold")).pack(pady=(14, 8))

        # ── Excel ──────────────────────────────────────────────────────
        self._row_with_button(
            label="Data file (.xlsx)",
            btn_text="Select…",
            cmd=self.on_select_excel,
            var=self.excel_path,
        )

        # ── Template ───────────────────────────────────────────────────
        self._row_with_button(
            label="Template (.docx)",
            btn_text="Select…",
            cmd=self.on_select_template,
            var=self.template_path,
        )

        # ── Mapping ────────────────────────────────────────────────────
        self._row_with_button(
            label="Mapping (.json)",
            btn_text="Select…",
            cmd=self.on_select_mapping,
            var=self.mapping_path,
        )

        ttk.Separator(self, orient="horizontal").pack(fill="x", padx=14, pady=10)

        # ── Line selector ──────────────────────────────────────────────
        line_frame = ttk.Frame(self)
        line_frame.pack(**pad)

        ttk.Label(line_frame, text="Line (row):", width=18, anchor="w").pack(side="left")

        ttk.Button(line_frame, text="−", width=3,
                   command=self._decrement_line).pack(side="left", padx=(0, 4))

        vcmd = (self.register(self._validate_int), "%P")
        self._line_entry = ttk.Entry(
            line_frame, textvariable=self.line_number,
            width=6, justify="center",
            validate="key", validatecommand=vcmd,
        )
        self._line_entry.pack(side="left")

        ttk.Button(line_frame, text="+", width=3,
                   command=self._increment_line).pack(side="left", padx=(4, 0))

        ttk.Separator(self, orient="horizontal").pack(fill="x", padx=14, pady=10)

        # ── Generate ───────────────────────────────────────────────────
        ttk.Button(
            self, text="Generate report",
            command=self.on_generate_report,
            style="Accent.TButton",
        ).pack(pady=(4, 16), ipadx=20, ipady=6)

    def _row_with_button(self, label, btn_text, cmd, var):
        """One row: label | button | path display."""
        frame = ttk.Frame(self)
        frame.pack(fill="x", padx=14, pady=3)

        ttk.Label(frame, text=label, width=18, anchor="w").pack(side="left")
        ttk.Button(frame, text=btn_text, command=cmd, width=10).pack(side="left", padx=(0, 8))
        ttk.Label(frame, textvariable=var, foreground="#1a5fb4",
                  wraplength=320, anchor="w").pack(side="left", fill="x", expand=True)

    # ------------------------------------------------------------------
    # Config persistence
    # ------------------------------------------------------------------

    def _restore_config(self):
        """Load last-used paths and line number from SQLite on startup."""
        excel = config_get(KEY_EXCEL)
        if excel:
            self.excel_path.set(excel)

        docx = config_get(KEY_DOCX)
        if docx:
            self.template_path.set(docx)

        mapping = config_get(KEY_MAPPING)
        self.mapping_path.set(mapping if mapping else DEFAULT_MAPPING)

        last_line = config_get(KEY_LINE)
        if last_line and last_line.isdigit():
            self.line_number.set(int(last_line) + 1)
        else:
            self.line_number.set(1)

    def _save_config(self, line: int):
        """Persist current UI state to SQLite."""
        config_set(KEY_EXCEL,   self.excel_path.get())
        config_set(KEY_DOCX,    self.template_path.get())
        config_set(KEY_MAPPING, self.mapping_path.get())
        config_set(KEY_LINE,    str(line))

    # ------------------------------------------------------------------
    # Line selector helpers
    # ------------------------------------------------------------------

    @staticmethod
    def _validate_int(value: str) -> bool:
        return value == "" or (value.isdigit() and int(value) >= 1)

    def _increment_line(self):
        self.line_number.set(self.line_number.get() + 1)

    def _decrement_line(self):
        current = self.line_number.get()
        if current > 1:
            self.line_number.set(current - 1)

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
            config_set(KEY_MAPPING, path)

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
        if line < 1:
            messagebox.showerror("Invalid line", "Line number must be ≥ 1.")
            return

        # Excel rows start at 2 (row 1 = header); the widget shows logical lines
        row_number = line + 1

        output_path = os.path.join(os.getcwd(), f"report_row_{row_number}.docx")

        try:
            generator = ReportGenerator(
                excel_path=excel,
                template_path=docx,
                mapping_path=mapping,
            )
            generator.generate(row_number=row_number, output_path=output_path)
        except Exception as exc:
            messagebox.showerror("Generation error", str(exc))
            return

        # Persist config after successful generation
        self._save_config(line)

        # Advance line selector
        self.line_number.set(line + 1)

        # Save report metadata to DB
        try:
            save_report(
                report_id=f"row_{row_number}_{datetime.datetime.now().strftime('%Y%m%d%H%M%S')}",
                created_at=datetime.datetime.now().isoformat(),
                excel_path=excel,
                template_path=docx,
                mapping_path=mapping,
                row_number=row_number,
                data={},
            )
        except Exception:
            pass  # DB logging failure must not block the user

        messagebox.showinfo("Done", f"Report generated:\n{output_path}")


if __name__ == "__main__":
    app = MainWindow()
    app.mainloop()
