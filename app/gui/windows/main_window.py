"""
Main application window — layout and event wiring only.
Business logic lives in report_actions.py and config_persistence.py.
"""
import os
import subprocess
import sys
import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog

from app.core.mapping_loader import MappingLoader
from app.core.state_manager import ReportStateManager
from app.core.excel_reader import ExcelReader
from app.core.processors import op_excel_day_counter
from app.core.profile_manager import ProfileManager, ProfileError
from app.gui.file_dialogs import (
    select_excel_file, select_docx_template,
    select_mapping_file, select_output_file,
)
from app.gui.line_selector import LineSelector
from app.gui.config_persistence import load_mapping_config
from app.gui.report_actions import resolve_output_path, generate_report
from app.gui.windows.generation_complete_dialog import GenerationCompleteDialog
from app.gui.windows.profile_dialogs import NewProfileDialog, UpdateProfileDialog
from app.gui.windows.guide_dialog import GuideDialog
from app.core.guide import build_profile_context
from app.utils.paths import get_resource_path

_STATE_PATH = os.path.abspath(
    os.path.join(os.path.dirname(__file__), "..", "report_state.json")
)


class MainWindow(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title("Report Generator")
        self.geometry("640x580")
        self.resizable(False, False)
        icon_path = get_resource_path("icon.ico")
        if os.path.exists(icon_path):
            self.iconbitmap(icon_path)

        # The three file paths always come from the active profile.
        self.excel_path    = tk.StringVar()
        self.template_path = tk.StringVar()
        self.mapping_path  = tk.StringVar()
        self.profile_var   = tk.StringVar()
        self.status_var    = tk.StringVar()
        self.output_dir    = tk.StringVar()
        self.report_name   = tk.StringVar()
        self.line_number   = tk.IntVar(value=2)

        self.profiles = ProfileManager()
        self._profile_slugs: dict[str, str] = {}   # display name -> slug
        self._current_slug: str | None = None
        self._issues = []

        self._auto_generated_report_name = False
        self._internal_report_name_update = False

        self._state_mgr = ReportStateManager(_STATE_PATH)

        self._build_ui()
        self._restore()

    # ------------------------------------------------------------------
    # UI construction
    # ------------------------------------------------------------------

    def _build_ui(self):
        ttk.Label(self, text="Report Generator", font=("Arial", 16, "bold")).pack(pady=(14, 8))

        self._profile_row()
        self._status_row()
        self._files_row()
        self._output_row()
        self._report_name_row()

        ttk.Separator(self, orient="horizontal").pack(fill="x", padx=14, pady=10)

        LineSelector(self, variable=self.line_number).pack(padx=14, pady=4)
        # Recompute report name and preview whenever line changes
        self.line_number.trace_add("write", lambda *_: self._on_line_number_changed())

        ttk.Separator(self, orient="horizontal").pack(fill="x", padx=14, pady=10)

        self._preview_label = ttk.Label(self, text="", foreground="#888888", font=("Arial", 9))
        self._preview_label.pack()

        self._generate_btn = ttk.Button(
            self, text="Generate report",
            command=self.on_generate_report,
        )
        self._generate_btn.pack(pady=(8, 10), ipadx=20, ipady=6)

        ttk.Separator(self, orient="horizontal").pack(fill="x", padx=14, pady=(0, 6))
        ttk.Button(self, text="📖 Guide — how the app, mappings, templates and data work",
                   command=self.on_open_guide).pack(pady=(0, 10))

        for i in range(9):
            self.bind_all(f"<Control-Key-{i + 1}>", lambda _e, i=i: self._select_by_index(i))
        self.bind_all("<F5>", lambda _e: self.on_refresh_profile())

        self._update_preview()

    def _profile_row(self):
        frame = ttk.Frame(self)
        frame.pack(fill="x", padx=14, pady=3)
        ttk.Label(frame, text="Profile", width=18, anchor="w").pack(side="left")
        self.profile_combobox = ttk.Combobox(
            frame, textvariable=self.profile_var, width=34, state="readonly",
        )
        self.profile_combobox.pack(side="left", fill="x", expand=True, padx=(0, 8))
        self.profile_combobox.bind("<<ComboboxSelected>>", lambda *_: self._on_profile_selected())
        ttk.Button(frame, text="New…", command=self.on_new_profile, width=8).pack(side="left")
        ttk.Button(frame, text="Update…", command=self.on_update_profile, width=9).pack(side="left", padx=(6, 0))

        self._profile_menu = tk.Menu(self, tearoff=0)
        for label, cmd in (
            ("Update profile…",         self.on_update_profile),
            ("Refresh (F5)",            self.on_refresh_profile),
            (None, None),
            ("Rename…",                 self.on_rename_profile),
            ("Duplicate…",              self.on_duplicate_profile),
            (None, None),
            ("Replace data file…",      self.on_replace_data),
            ("Replace template…",       self.on_replace_template),
            ("Replace mapping…",        self.on_replace_mapping),
            (None, None),
            ("Open profile folder",     self.on_open_profile_folder),
            ("Export profile (.zip)…",  self.on_export_profile),
            ("Import profile (.zip)…",  self.on_import_profile),
            (None, None),
            ("Delete profile…",         self.on_delete_profile),
        ):
            if label is None:
                self._profile_menu.add_separator()
            else:
                self._profile_menu.add_command(label=label, command=cmd)
        self._menu_btn = ttk.Button(frame, text="⋯", width=3, command=self._show_profile_menu)
        self._menu_btn.pack(side="left", padx=(6, 0))

    def _show_profile_menu(self):
        x = self._menu_btn.winfo_rootx()
        y = self._menu_btn.winfo_rooty() + self._menu_btn.winfo_height()
        self._profile_menu.tk_popup(x, y)

    def _status_row(self):
        frame = ttk.Frame(self)
        frame.pack(fill="x", padx=14, pady=(0, 3))
        ttk.Label(frame, text="", width=18).pack(side="left")
        self._status_label = ttk.Label(frame, textvariable=self.status_var, anchor="w")
        self._status_label.pack(side="left", fill="x", expand=True)
        self._details_btn = ttk.Button(frame, text="Details…", width=10, command=self.on_show_issues)
        self._details_btn.pack(side="right")
        ttk.Button(frame, text="⟳ Refresh", width=10, command=self.on_refresh_profile).pack(side="right", padx=(0, 6))

    def _files_row(self):
        frame = ttk.Frame(self)
        frame.pack(fill="x", padx=14, pady=3)
        ttk.Label(frame, text="Profile files", width=18, anchor="w").pack(side="left")
        for text, var in (("📂 Mapping", self.mapping_path),
                          ("📂 Template", self.template_path),
                          ("📂 Data", self.excel_path)):
            ttk.Button(frame, text=text, width=12,
                       command=lambda v=var: self._open_file(v.get())).pack(side="left", padx=(0, 6))

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
        if self._internal_report_name_update:
            self._internal_report_name_update = False
            return

        self._auto_generated_report_name = False
        mapping = self.mapping_path.get()
        if not mapping:
            self._update_preview()
            return
        try:
            MappingLoader(mapping).update_file_name_field(name=self.report_name.get())
        except Exception:
            pass
        self._update_preview()

    def _on_line_number_changed(self):
        try:
            line = self.line_number.get()
        except tk.TclError:          # entry is momentarily empty while typing
            return
        if self._current_slug and line >= 2:
            try:
                self.profiles.set_next_line(self._current_slug, line)
            except Exception:
                pass
        if self._auto_generated_report_name:
            mapping = self.mapping_path.get()
            if mapping:
                new_name = self._resolve_mapping_name(mapping)
                if new_name:
                    self._set_report_name(new_name)
        self._update_preview()

    def _set_report_name(self, name: str):
        self._internal_report_name_update = True
        self.report_name.set(name)

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
            name = self.report_name.get().strip()
            if numero:
                if name and name.endswith(f" {numero}"):
                    display = name
                else:
                    display = f"{name} {numero}".strip() if name else numero
            else:
                display = "—"
            self._preview_label.config(text=f"Report number: {display}")
        except Exception:
            pass

    # ------------------------------------------------------------------
    # Profiles
    # ------------------------------------------------------------------

    def _restore(self):
        migration = None
        try:
            migration = self.profiles.migrate_legacy()
        except Exception:
            migration = None

        self._refresh_profiles()
        slug = self.profiles.active_slug()
        if not slug and self._profile_slugs:
            slug = next(iter(self._profile_slugs.values()))
        if slug:
            self._activate(slug)
        else:
            self._show_no_profile()

        if migration and migration.get("created"):
            names = "\n".join(f"  • {n}" for n in migration["created"])
            self.after(200, lambda: messagebox.showinfo(
                "Profiles created",
                "Your existing mappings were turned into profiles "
                "(their files were copied; the originals were not touched):\n\n" + names,
            ))

    def _refresh_profiles(self):
        profiles = self.profiles.list()
        self._profile_slugs = {p.name: p.slug for p in profiles}
        self.profile_combobox["values"] = list(self._profile_slugs)

    def _select_by_index(self, index: int):
        slugs = list(self._profile_slugs.values())
        if 0 <= index < len(slugs):
            self._activate(slugs[index])

    def _on_profile_selected(self):
        slug = self._profile_slugs.get(self.profile_var.get())
        if slug:
            self._activate(slug)

    def _activate(self, slug: str):
        """Switch the whole window to a profile: files, output dir, row, report name."""
        try:
            profile = self.profiles.get(slug)
        except ProfileError:
            self._show_no_profile()
            return

        self._current_slug = slug
        self.profiles.set_active(slug)
        self.profile_var.set(profile.name)
        self.mapping_path.set(profile.mapping_path)
        self.excel_path.set(profile.data_path)
        self.template_path.set(profile.template_path)

        cfg = load_mapping_config(profile.mapping_path)
        self.output_dir.set(profile.output_dir or cfg.get("output_dir", "") or "")
        self.line_number.set(profile.next_line)

        self._apply_mapping_config(profile.mapping_path)
        self._refresh_status()
        self._update_preview()

    def _show_no_profile(self):
        self._current_slug = None
        for var in (self.profile_var, self.mapping_path, self.excel_path,
                    self.template_path, self.output_dir, self.report_name):
            var.set("")
        self._issues = []
        self.status_var.set("No profile yet — click New… to create one.")
        self._status_label.configure(foreground="#666666")
        self._details_btn.configure(state="disabled")
        self._generate_btn.configure(state="disabled")
        self._update_preview()

    def _refresh_status(self):
        slug = self._current_slug
        if not slug:
            return
        self.configure(cursor="watch")
        self.update_idletasks()
        try:
            self._issues = self.profiles.validate(slug)
        except Exception as exc:
            from app.core.profile_manager import Issue
            self._issues = [Issue("error", f"Could not check the profile: {exc}")]
        finally:
            self.configure(cursor="")

        errors   = [i for i in self._issues if i.severity == "error"]
        warnings = [i for i in self._issues if i.severity == "warning"]
        if errors:
            more = f"  (+{len(errors) - 1} more)" if len(errors) > 1 else ""
            self.status_var.set(f"✖ {errors[0].message}{more}")
            color = "#c01c28"
        elif warnings:
            self.status_var.set(f"⚠ {len(warnings)} warning(s) — the report can still be generated")
            color = "#b5651d"
        else:
            self.status_var.set("✔ Ready")
            color = "#26a269"
        self._status_label.configure(foreground=color)
        self._details_btn.configure(state="normal" if self._issues else "disabled")
        self._generate_btn.configure(state="disabled" if errors else "normal")

    def on_show_issues(self):
        if not self._issues:
            return
        lines = [("✖ " if i.severity == "error" else "⚠ ") + i.message for i in self._issues]
        messagebox.showwarning("Profile check", "\n".join(lines))

    def _require_profile(self) -> str | None:
        if not self._current_slug:
            messagebox.showinfo("Profile", "Create or select a profile first.")
            return None
        return self._current_slug

    def on_open_guide(self):
        """Open the guide (one window at a time); it can copy itself + the active profile for an AI."""
        existing = getattr(self, "_guide_window", None)
        try:
            if existing is not None and existing.winfo_exists():
                existing.lift()
                existing.focus_set()
                return
        except tk.TclError:
            pass
        self._guide_window = GuideDialog(self, context_provider=self._profile_context)

    def _profile_context(self):
        """Markdown description of the active profile (fresh check, headers only — no cell values)."""
        slug = self._current_slug
        if not slug:
            return None
        return build_profile_context(self.profiles.get(slug), self.profiles.validate(slug))

    def on_refresh_profile(self):
        """Re-read the profile's files from disk (mapping, template, data) and re-check them."""
        self._refresh_profiles()                       # also picks up folders added/removed by hand
        slugs = list(self._profile_slugs.values())
        if self._current_slug in slugs:
            self._activate(self._current_slug)
        elif slugs:
            self._activate(slugs[0])
        else:
            self._show_no_profile()

    def on_update_profile(self):
        """Same window as 'New…', for the current profile: rename it and/or replace its files."""
        slug = self._require_profile()
        if not slug:
            return
        profile = self.profiles.get(slug)
        dialog = UpdateProfileDialog(self, profile.name, self.profiles.file_info(slug))
        if not dialog.result:
            return
        r = dialog.result
        done = []
        try:
            if r["name"] != profile.name:
                self.profiles.rename(slug, r["name"])
                done.append("name")
            if r["mapping"]:
                self.profiles.replace_mapping(slug, r["mapping"])
                done.append("mapping")
            if r["template"]:
                self.profiles.replace_template(slug, r["template"])
                done.append("template")
            if r["data"]:
                self.profiles.replace_data(slug, r["data"])
                done.append("data")
        except (ProfileError, OSError, ValueError) as exc:
            extra = f"\n\nAlready applied: {', '.join(done)}." if done else ""
            messagebox.showerror("Update profile", f"{exc}{extra}")
        self._refresh_profiles()
        self._activate(slug)

    # ---- create / rename / duplicate / delete ----------------------------

    def on_new_profile(self):
        existing = [(slug, name) for name, slug in self._profile_slugs.items()]
        dialog = NewProfileDialog(self, existing)
        if not dialog.result:
            return
        r = dialog.result
        try:
            profile = self.profiles.create_from(
                r["name"], source_slug=r["source_slug"],
                mapping=r["mapping"], template=r["template"], data=r["data"],
            )
        except (ProfileError, OSError, ValueError) as exc:
            messagebox.showerror("New profile", str(exc))
            return
        self._refresh_profiles()
        self._activate(profile.slug)

    def on_rename_profile(self):
        slug = self._require_profile()
        if not slug:
            return
        current = self.profiles.get(slug).name
        new_name = simpledialog.askstring("Rename profile", "New name:",
                                          initialvalue=current, parent=self)
        if not new_name or new_name.strip() == current:
            return
        try:
            self.profiles.rename(slug, new_name)
        except ProfileError as exc:
            messagebox.showerror("Rename profile", str(exc))
            return
        self._refresh_profiles()
        self.profile_var.set(self.profiles.get(slug).name)

    def on_duplicate_profile(self):
        slug = self._require_profile()
        if not slug:
            return
        current = self.profiles.get(slug).name
        new_name = simpledialog.askstring("Duplicate profile", "Name of the copy:",
                                          initialvalue=f"{current} copy", parent=self)
        if not new_name:
            return
        try:
            profile = self.profiles.duplicate(slug, new_name)
        except (ProfileError, OSError) as exc:
            messagebox.showerror("Duplicate profile", str(exc))
            return
        self._refresh_profiles()
        self._activate(profile.slug)

    def on_delete_profile(self):
        slug = self._require_profile()
        if not slug:
            return
        name = self.profiles.get(slug).name
        if not messagebox.askyesno(
            "Delete profile",
            f"Delete the profile '{name}'?\n\n"
            "Its folder (mapping, template and data) is moved to the "
            "'_trash' folder inside the profiles folder, so it can be recovered.",
            icon="warning",
        ):
            return
        try:
            self.profiles.delete(slug)
        except (ProfileError, OSError) as exc:
            messagebox.showerror("Delete profile", str(exc))
            return
        self._refresh_profiles()
        remaining = list(self._profile_slugs.values())
        if remaining:
            self._activate(remaining[0])
        else:
            self._show_no_profile()

    # ---- replace files / folder / export / import ---------------------------

    def _replace(self, picker, replacer, label: str):
        slug = self._require_profile()
        if not slug:
            return
        path = picker()
        if not path:
            return
        try:
            replacer(slug, path)
        except (ProfileError, OSError, ValueError) as exc:
            messagebox.showerror(label, str(exc))
            return
        self._refresh_status()
        self._apply_mapping_config(self.mapping_path.get())
        self._update_preview()

    def on_replace_data(self):
        self._replace(select_excel_file, self.profiles.replace_data, "Replace data file")

    def on_replace_template(self):
        self._replace(select_docx_template, self.profiles.replace_template, "Replace template")

    def on_replace_mapping(self):
        self._replace(select_mapping_file, self.profiles.replace_mapping, "Replace mapping")

    def on_open_profile_folder(self):
        slug = self._require_profile()
        if slug:
            self._open_file(str(self.profiles.get(slug).folder))

    def on_export_profile(self):
        slug = self._require_profile()
        if not slug:
            return
        dest = filedialog.asksaveasfilename(
            title="Export profile", defaultextension=".zip",
            initialfile=f"{slug}.zip", filetypes=[("Profile archive", "*.zip")],
        )
        if not dest:
            return
        try:
            self.profiles.export_zip(slug, dest)
        except (ProfileError, OSError) as exc:
            messagebox.showerror("Export profile", str(exc))
            return
        messagebox.showinfo("Export profile", f"Profile exported to:\n{dest}")

    def on_import_profile(self):
        src = filedialog.askopenfilename(
            title="Import profile", filetypes=[("Profile archive", "*.zip")],
        )
        if not src:
            return
        try:
            profile = self.profiles.import_zip(src)
        except (ProfileError, OSError) as exc:
            messagebox.showerror("Import profile", str(exc))
            return
        self._refresh_profiles()
        self._activate(profile.slug)

    # ------------------------------------------------------------------
    # Other button callbacks
    # ------------------------------------------------------------------

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
            if self._current_slug:
                self.profiles.set_output_dir(self._current_slug, directory)

    def _apply_mapping_config(self, mapping_path: str):
        # Load report name from file_name.name or from the resolved mapping file_name.
        try:
            fn = MappingLoader(mapping_path).load_file_name_field()
            name = fn.get("name", "")
            resolved_name = ""
            if not name or ("{" in name and "}" in name):
                resolved_name = self._resolve_mapping_name(mapping_path)

            if resolved_name:
                self._auto_generated_report_name = True
                self._set_report_name(resolved_name)
            elif name:
                self._auto_generated_report_name = False
                self._set_report_name(name)
            else:
                self._auto_generated_report_name = False
                self._set_report_name("")
        except Exception:
            self._auto_generated_report_name = False

    def _resolve_mapping_name(self, mapping_path: str) -> str:
        excel_path = self.excel_path.get()
        template_path = self.template_path.get()
        row_number = self.line_number.get()
        if not excel_path or not template_path or row_number < 2:
            return ""
        try:
            _, filename = resolve_output_path(
                mapping_path, excel_path, template_path, row_number,
                profile_slug=self._current_slug,
            )
            if filename.lower().endswith(".docx"):
                filename = filename[:-5]
            return filename
        except Exception:
            return ""

    def _update_mapping_config(self, **kwargs):
        mapping = self.mapping_path.get()
        if not mapping:
            return
        try:
            MappingLoader(mapping).update_config(**kwargs)
        except Exception:
            pass

    def on_generate_report(self):
        slug    = self._current_slug
        excel   = self.excel_path.get()
        docx    = self.template_path.get()
        mapping = self.mapping_path.get()
        line    = self.line_number.get()

        if not slug:
            messagebox.showerror("No profile", "Create or select a profile first.")
            return
        if any(i.severity == "error" for i in self._issues):
            messagebox.showerror("Profile problem", "Fix the problems shown under the profile first.")
            return
        if line < 2:
            messagebox.showerror("Invalid line", "Row number must be ≥ 2 (row 1 is the header).")
            return

        chosen_output_dir = self.output_dir.get().strip() or None
        suggested_dir, suggested_name = resolve_output_path(
            mapping, excel, docx, line, profile_slug=slug, output_dir=chosen_output_dir)
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
            self.profiles.set_output_dir(slug, final_dir)
            self.output_dir.set(final_dir)

        try:
            final_path = generate_report(excel, docx, mapping, row_number=line,
                                         output_path=output_path, profile_slug=slug)
        except Exception as exc:
            messagebox.showerror("Generation error", str(exc))
            return

        self.profiles.set_next_line(slug, line + 1)
        self.line_number.set(line + 1)
        self._update_preview()

        GenerationCompleteDialog(self, final_path)


if __name__ == "__main__":
    app = MainWindow()
    app.mainloop()
