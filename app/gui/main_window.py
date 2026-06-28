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
from app.gui.file_dialogs import (
    select_excel_file, select_docx_template,
    select_mapping_file, select_output_file,
)
from app.gui.line_selector import LineSelector
from app.gui.config_persistence import (
    restore_config, save_config, save_mapping_path, load_mapping_config, save_output_dir,
)
from app.gui.report_actions import resolve_output_path, generate_report

DEFAULT_MAPPING = os.path.abspath(
    os.path.join(os.path.dirname(__file__), "..", "mappings", "data.json")
)

# Dummy state_path (only needed for legacy migration)
_STATE_PATH = os.path.abspath(
    os.path.join(os.path.dirname(__file__), "..", "report_state.json")
)


class MainWindow(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title("Report Generator")
        self.geometry("620x540")
        self.resizable(False, False)

        self.excel_path    = tk.StringVar()
        self.template_path = tk.StringVar()
        self.mapping_path  = tk.StringVar()
        self.output_dir    = tk.StringVar()
        self.report_name   = tk.StringVar()
        self.line_number   = tk.IntVar(value=2)

        # Report counter — editable by the user
        self._state_mgr = ReportStateManager(_STATE_PATH)
        self.report_counter = tk.IntVar(value=self._state_mgr.get_current_counter())

        self._build_ui()
        self._restore()

    # ------------------------------------------------------------------
    # UI construction
    # ------------------------------------------------------------------

    def _build_ui(self):
        ttk.Label(self, text="Report Generator", font=("Arial", 16, "bold")).pack(pady=(14, 8))

        self._path_row("Mapping (.json)",   "Select…", self.on_select_mapping,  self.mapping_path)
        self._path_row("Data file (.xlsx)", "Select…", self.on_select_excel,    self.excel_path)
        self._path_row("Template (.docx)",  "Select…", self.on_select_template, self.template_path)
        self._output_row()
        self._report_name_row()

        ttk.Separator(self, orient="horizontal").pack(fill="x", padx=14, pady=10)

        LineSelector(self, variable=self.line_number).pack(padx=14, pady=4)

        # Report counter row
        self._counter_row()

        ttk.Separator(self, orient="horizontal").pack(fill="x", padx=14, pady=10)

        # Preview of the next report number
        self._preview_label = ttk.Label(self, text="", foreground="#888888", font=("Arial", 9))
        self._preview_label.pack()
        self._update_preview()

        ttk.Button(
            self, text="Generate report",
            command=self.on_generate_report,
        ).pack(pady=(8, 16), ipadx=20, ipady=6)

    def _path_row(self, label: str, btn_text: str, cmd, var: tk.StringVar):
        frame = ttk.Frame(self)
        frame.pack(fill="x", padx=14, pady=3)
        ttk.Label(frame, text=label, width=18, anchor="w").pack(side="left")
        ttk.Button(frame, text=btn_text, command=cmd, width=10).pack(side="left", padx=(0, 8))
        ttk.Label(frame, textvariable=var, foreground="#1a5fb4",
                  wraplength=320, anchor="w").pack(side="left", fill="x", expand=True)

    def _output_row(self):
        frame = ttk.Frame(self)
        frame.pack(fill="x", padx=14, pady=3)
        ttk.Label(frame, text="Output Location", width=18, anchor="w").pack(side="left")
        ttk.Button(frame, text="Select…", command=self.on_select_output_dir, width=10).pack(side="left", padx=(0, 4))
        ttk.Button(frame, text="📂 Open", command=self.on_open_output_dir, width=8).pack(side="left", padx=(0, 8))
        ttk.Label(frame, textvariable=self.output_dir, foreground="#1a5fb4",
                  wraplength=260, anchor="w").pack(side="left", fill="x", expand=True)

    def _counter_row(self):
        frame = ttk.Frame(self)
        frame.pack(fill="x", padx=14, pady=3)
        ttk.Label(frame, text="Report counter:", width=18, anchor="w").pack(side="left")

        vcmd = (self.register(self._validate_counter), "%P")
        ttk.Entry(
            frame, textvariable=self.report_counter,
            width=6, justify="center",
            validate="key", validatecommand=vcmd,
        ).pack(side="left", padx=(0, 8))

        ttk.Label(
            frame,
            text="(next report will use counter + 1)",
            foreground="#888888", font=("Arial", 9),
        ).pack(side="left")

        # Update preview whenever counter changes
        self.report_counter.trace_add("write", lambda *_: self._update_preview())

    def _report_name_row(self):
        frame = ttk.Frame(self)
        frame.pack(fill="x", padx=14, pady=3)
        ttk.Label(frame, text="Name of the report:", width=18, anchor="w").pack(side="left")
        entry = ttk.Entry(frame, textvariable=self.report_name, width=30)
        entry.pack(side="left", padx=(0, 8))
        ttk.Label(
            frame, text="(prefix used in report number)",
            foreground="#888888", font=("Arial", 9),
        ).pack(side="left")
        # Persist to mapping config whenever the user edits the field
        self.report_name.trace_add("write", self._on_report_name_changed)

    def _on_report_name_changed(self, *_):
        name = self.report_name.get()
        self._update_mapping_config(report_prefix=name)
        self._update_preview()

    @staticmethod
    def _validate_counter(value: str) -> bool:
        return value == "" or (value.isdigit() and int(value) >= 0)

    def _update_preview(self):
        try:
            prefix = self._state_mgr.get_today_prefix()
            next_n = self.report_counter.get() + 1
            name   = self.report_name.get().strip()
            number = f"{prefix}-{next_n:02d}"
            display = f"{name} {number}".strip() if name else number
            self._preview_label.config(text=f"Next report number: {display}")
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

        # Restore report name (prefix) from mapping config
        report_prefix = state.get("mapping_cfg", {}).get("report_prefix", "")
        if report_prefix:
            self.report_name.set(report_prefix)

        # Refresh counter from DB in case another session changed it
        self.report_counter.set(self._state_mgr.get_current_counter())

    # ------------------------------------------------------------------
    # Button callbacks
    # ------------------------------------------------------------------

    def on_select_excel(self):
        path = select_excel_file()
        if path:
            self.excel_path.set(path)
            self._update_mapping_config(data_file=path)

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
        if cfg.get("report_prefix"):
            self.report_name.set(cfg["report_prefix"])

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

        # Persist the manually-set counter before generation (so the generator picks up N+1)
        try:
            self._state_mgr.set_counter(self.report_counter.get())
        except Exception:
            pass

        # Determine output dir: prefer the UI field, then mapping config, then cwd
        chosen_output_dir = self.output_dir.get().strip() or None

        # Pre-compute the suggested path from mapping config + file_name field
        suggested_dir, suggested_name = resolve_output_path(mapping, excel, docx, line)
        if chosen_output_dir:
            suggested_dir = chosen_output_dir

        # Ask the user to confirm or change the save location
        output_path = select_output_file(
            initial_dir=suggested_dir,
            initial_file=suggested_name,
        )
        if not output_path:
            return  # user cancelled

        # If the user changed the output folder, persist it
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

        # Refresh counter display after generation (it was incremented inside generator)
        self.report_counter.set(self._state_mgr.get_current_counter())
        self._update_preview()

        messagebox.showinfo("Done", f"Report generated:\n{final_path}")


if __name__ == "__main__":
    app = MainWindow()
    app.mainloop()