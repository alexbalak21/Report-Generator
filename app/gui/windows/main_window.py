"""
Main application window — layout and event wiring only.
Business logic lives in report_actions.py and config_persistence.py.
"""
import os
import subprocess
import sys
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

from app.core.mapping_loader import MappingLoader
from app.core.state_manager import ReportStateManager
from app.core.excel_reader import ExcelReader
from app.core.processors import op_excel_day_counter
from app.gui.file_dialogs import (
    select_excel_file, select_docx_template,
    select_mapping_file, select_output_file,
)
from app.gui.line_selector import LineSelector
from app.gui.config_persistence import (
    restore_config, save_config, save_mapping_path, load_mapping_config, save_output_dir,
)
from app.gui.report_actions import resolve_output_path, generate_report
from app.gui.windows.generation_complete_dialog import GenerationCompleteDialog

DEFAULT_MAPPING = os.path.abspath(
    os.path.join(os.path.dirname(__file__), "..", "mappings", "data.json")
)

_STATE_PATH = os.path.abspath(
    os.path.join(os.path.dirname(__file__), "..", "report_state.json")
)


class MainWindow(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title("Report Generator")
        self.geometry("620x500")
        self.resizable(False, False)

        self.excel_path    = tk.StringVar()
        self.template_path = tk.StringVar()
        self.mapping_path  = tk.StringVar()
        self.output_dir    = tk.StringVar()
        self.report_name   = tk.StringVar()
        self.line_number   = tk.IntVar(value=2)

        self._state_mgr = ReportStateManager(_STATE_PATH)

        self._build_ui()
        self._restore()

    # ------------------------------------------------------------------
    # UI construction
    # ------------------------------------------------------------------

    def _build_ui(self):
        ttk.Label(self, text="Report Generator", font=("Arial", 16, "bold")).pack(pady=(14, 8))

        self._path_row(
            "Mapping file", "Select…", self.on_select_mapping, self.mapping_path,
            open_cmd=lambda: self._open_file(self.mapping_path.get()),
        )
        self._path_row(
            "Data file (.xlsx)", "Select…", self.on_select_excel, self.excel_path,
            open_cmd=lambda: self._open_file(self.excel_path.get()),
        )
        self._path_row(
            "Template (.docx)", "Select…", self.on_select_template, self.template_path,
            open_cmd=lambda: self._open_file(self.template_path.get()),
        )
        self._output_row()
        self._report_name_row()

        ttk.Separator(self, orient="horizontal").pack(fill="x", padx=14, pady=10)

        LineSelector(self, variable=self.line_number).pack(padx=14, pady=4)
        # Recompute preview whenever line changes
        self.line_number.trace_add("write", lambda *_: self._update_preview())

        ttk.Separator(self, orient="horizontal").pack(fill="x", padx=14, pady=10)

        self._preview_label = ttk.Label(self, text="", foreground="#888888", font=("Arial", 9))
        self._preview_label.pack()

        ttk.Button(
            self, text="Generate report",
            command=self.on_generate_report,
        ).pack(pady=(8, 16), ipadx=20, ipady=6)

        self._update_preview()

    def _path_row(self, label: str, btn_text: str, cmd, var: tk.StringVar, open_cmd=None):
        frame = ttk.Frame(self)
        frame.pack(fill="x", padx=14, pady=3)
        ttk.Label(frame, text=label, width=18, anchor="w").pack(side="left")
        ttk.Button(frame, text=btn_text, command=cmd, width=10).pack(side="left", padx=(0, 8))
        ttk.Label(frame, textvariable=var, foreground="#1a5fb4",
                  wraplength=260, anchor="w").pack(side="left", fill="x", expand=True)
        ttk.Button(frame, text="📂 Open", command=open_cmd, width=10,
                   state="normal" if open_cmd else "disabled").pack(side="left", padx=(8, 0))

    def _output_row(self):
        frame = ttk.Frame(self)
        frame.pack(fill="x", padx=14, pady=3)
        ttk.Label(frame, text="Output Location", width=18, anchor="w").pack(side="left")
        ttk.Button(frame, text="Select…", command=self.on_select_output_dir, width=10).pack(side="left", padx=(0, 4))
        ttk.Label(frame, textvariable=self.output_dir, foreground="#1a5fb4",
                  wraplength=260, anchor="w").pack(side="left", fill="x", expand=True)
        ttk.Button(frame, text="📂 Open", command=self.on_open_output_dir, width=10).pack(side="left", padx=(8, 0))

    def _report_name_row(self):
        frame = ttk.Frame(self)
        frame.pack(fill="x", padx=14, pady=3)
        ttk.Label(frame, text="Name of the report:", width=18, anchor="w").pack(side="left")
        ttk.Entry(frame, textvariable=self.report_name, width=36).pack(side="left", padx=(0, 8))
        self.report_name.trace_add("write", self._on_report_name_changed)

    def _open_file(self, path: str):
        if not path or not os.path.exists(path):
            messagebox.showwarning("Open file", "The selected file does not exist.")
            return
        if sys.platform == "win32":
            os.startfile(path)
        elif sys.platform == "darwin":
            subprocess.Popen(["open", path])
        else:
            subprocess.Popen(["xdg-open", path])

    def _on_report_name_changed(self, *_):
        mapping = self.mapping_path.get()
        if not mapping:
            return
        try:
            MappingLoader(mapping).update_file_name_field(name=self.report_name.get())
        except Exception:
            pass
        self._update_preview()

    # ------------------------------------------------------------------
    # Preview — computed from Excel using the selected row
    # ------------------------------------------------------------------

    def _compute_numero_rapport(self) -> str:
        """
        Read the 'date rapport' column for the selected row and compute
        the Excel-based day counter (yymmdd-N), exactly as the generator does.
        Returns a string like '260621-3' or '' on any error.
        """
        excel_path   = self.excel_path.get()
        mapping_path = self.mapping_path.get()
        row_number   = self.line_number.get()

        if not excel_path or not mapping_path or row_number < 2:
            return ""

        try:
            loader  = MappingLoader(mapping_path)
            rules   = loader.load()
            cfg     = loader.load_config()
            date_fmt = cfg.get("date_format", "%d/%m/%Y")

            # Accept either the legacy `numero_rapport` key or the new `sample_number` key
            numero_rule = rules.get("sample_number") or rules.get("numero_rapport")
            if not isinstance(numero_rule, dict):
                return ""

            reader = ExcelReader(excel_path)
            reader.load()
            raw_row  = reader.get_row_as_dict(row_number)

            row_data = {}
            import datetime as dt
            for k, v in raw_row.items():
                if isinstance(v, (dt.date, dt.datetime)):
                    row_data[k] = v.strftime(date_fmt)
                elif v is None:
                    row_data[k] = ""
                else:
                    row_data[k] = str(v).strip()

            result = op_excel_day_counter(numero_rule, row_data, reader, row_number)
            return result or ""
        except Exception:
            return ""

    def _update_preview(self):
        try:
            numero = self._compute_numero_rapport()
            name   = self.report_name.get().strip()
            if numero:
                display = f"{name} {numero}".strip() if name else numero
            else:
                display = "—"
            self._preview_label.config(text=f"Report number: {display}")
        except Exception:
            pass

    # ------------------------------------------------------------------
    # Startup restore
    # ------------------------------------------------------------------

    def _restore(self):
        state = restore_config()
        self.line_number.set(state["line"])

        mapping = state["mapping"] or DEFAULT_MAPPING
        self.mapping_path.set(mapping)
        self._apply_mapping_config(mapping)

        if state["excel"]:
            self.excel_path.set(state["excel"])
        if state["docx"]:
            self.template_path.set(state["docx"])
        if state.get("output_dir"):
            self.output_dir.set(state["output_dir"])

        self._update_preview()

    # ------------------------------------------------------------------
    # Button callbacks
    # ------------------------------------------------------------------

    def on_select_excel(self):
        path = select_excel_file()
        if path:
            self.excel_path.set(path)
            self._update_mapping_config(data_file=path)
            self._update_preview()

    def on_select_template(self):
        path = select_docx_template()
        if path:
            self.template_path.set(path)
            self._update_mapping_config(template_file=path)

    def on_select_mapping(self):
        path = select_mapping_file()
        if path:
            self.mapping_path.set(path)
            save_mapping_path(path)
            self._apply_mapping_config(path)
            self._update_preview()

    def on_open_output_dir(self):
        directory = self.output_dir.get().strip()
        if not directory or not os.path.isdir(directory):
            messagebox.showwarning("Output folder", "No valid output folder selected yet.")
            return
        if sys.platform == "win32":
            os.startfile(directory)
        elif sys.platform == "darwin":
            subprocess.Popen(["open", directory])
        else:
            subprocess.Popen(["xdg-open", directory])

    def on_select_output_dir(self):
        directory = filedialog.askdirectory(title="Select output folder")
        if directory:
            self.output_dir.set(directory)
            self._update_mapping_config(output_dir=directory)
            save_output_dir(directory)

    def _apply_mapping_config(self, mapping_path: str):
        cfg = load_mapping_config(mapping_path)
        if cfg.get("data_file"):
            self.excel_path.set(cfg["data_file"])
        if cfg.get("template_file"):
            self.template_path.set(cfg["template_file"])
        if cfg.get("output_dir"):
            self.output_dir.set(cfg["output_dir"])

        # Load report name from file_name.name
        try:
            fn = MappingLoader(mapping_path).load_file_name_field()
            name = fn.get("name", "")
            if name:
                self.report_name.set(name)
        except Exception:
            pass

    def _update_mapping_config(self, **kwargs):
        mapping = self.mapping_path.get()
        if not mapping:
            return
        try:
            MappingLoader(mapping).update_config(**kwargs)
        except Exception:
            pass

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

        chosen_output_dir = self.output_dir.get().strip() or None
        suggested_dir, suggested_name = resolve_output_path(mapping, excel, docx, line)
        if chosen_output_dir:
            suggested_dir = chosen_output_dir

        output_path = select_output_file(
            initial_dir=suggested_dir,
            initial_file=suggested_name,
        )
        if not output_path:
            return

        final_dir = os.path.dirname(os.path.abspath(output_path))
        if final_dir != os.path.abspath(suggested_dir):
            self._update_mapping_config(output_dir=final_dir)
            save_output_dir(final_dir)
            self.output_dir.set(final_dir)

        try:
            final_path = generate_report(excel, docx, mapping, row_number=line,
                                         output_path=output_path)
        except Exception as exc:
            messagebox.showerror("Generation error", str(exc))
            return

        save_config(excel, docx, mapping, line)
        self.line_number.set(line + 1)
        self._update_preview()

        GenerationCompleteDialog(self, final_path)


if __name__ == "__main__":
    app = MainWindow()
    app.mainloop()