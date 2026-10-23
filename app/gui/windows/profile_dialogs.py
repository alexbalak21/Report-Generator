"""
Dialogs used by the profile selector (kept apart from main_window.py).
"""
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from app.core.profile_manager import MAPPING_FILE, TEMPLATE_FILE, DATA_FILE


class ProfileDialog(tk.Toplevel):
    """
    One window for both "New profile" and "Update profile".

    After the dialog closes, `self.result` is None (cancelled) or a dict:
        {"name", "mapping", "template", "data", "source_slug"}
    Empty file fields mean:  new profile -> no file yet,  update -> keep the current file.
    """

    _NO_SOURCE = "(empty profile)"

    def __init__(self, parent, existing: list[tuple[str, str]],
                 current_name: str | None = None, file_info: dict | None = None):
        """
        existing     — [(slug, display name), …] of profiles ("start from" list, new mode only)
        current_name — set for update mode: the profile's present name
        file_info    — update mode: ProfileManager.file_info() of the profile
        """
        super().__init__(parent)
        self.update_mode = current_name is not None
        self.title("Update profile" if self.update_mode else "New profile")
        self.resizable(False, False)
        self.transient(parent)
        self.result = None
        self._existing = dict((name, slug) for slug, name in existing)
        self._file_info = file_info or {}

        self.name_var     = tk.StringVar(value=current_name or "")
        self.source_var   = tk.StringVar(value=self._NO_SOURCE)
        self.mapping_var  = tk.StringVar()
        self.template_var = tk.StringVar()
        self.data_var     = tk.StringVar()

        body = ttk.Frame(self, padding=14)
        body.pack(fill="both", expand=True)

        ttk.Label(body, text="Profile name", width=16, anchor="w").grid(row=0, column=0, sticky="w", pady=3)
        name_entry = ttk.Entry(body, textvariable=self.name_var, width=44)
        name_entry.grid(row=0, column=1, columnspan=2, sticky="we", pady=3)

        if not self.update_mode:
            ttk.Label(body, text="Start from", width=16, anchor="w").grid(row=1, column=0, sticky="w", pady=3)
            ttk.Combobox(
                body, textvariable=self.source_var, state="readonly", width=41,
                values=[self._NO_SOURCE, *self._existing.keys()],
            ).grid(row=1, column=1, columnspan=2, sticky="we", pady=3)

        ttk.Separator(body).grid(row=2, column=0, columnspan=3, sticky="we", pady=8)
        hint = ("Choose a file only to replace it — empty fields keep the current file "
                "(the old one is saved as .bak)"
                if self.update_mode else "Files (optional — copied into the profile)")
        ttk.Label(body, text=hint, foreground="#666666").grid(row=3, column=0, columnspan=3, sticky="w")

        self._file_row(body, 4, "Mapping (.json)", self.mapping_var, MAPPING_FILE,
                       [("Mapping files", "*.json *.xlsx")])
        self._file_row(body, 6, "Template (.docx)", self.template_var, TEMPLATE_FILE,
                       [("Word documents", "*.docx")])
        self._file_row(body, 8, "Data (.xlsx)", self.data_var, DATA_FILE,
                       [("Excel files", "*.xlsx *.xlsm")])

        buttons = ttk.Frame(body)
        buttons.grid(row=10, column=0, columnspan=3, sticky="e", pady=(12, 0))
        ttk.Button(buttons, text="Cancel", command=self.destroy).pack(side="right")
        ttk.Button(buttons, text="Save" if self.update_mode else "Create",
                   command=self._on_ok).pack(side="right", padx=(0, 8))

        self.bind("<Return>", lambda *_: self._on_ok())
        self.bind("<Escape>", lambda *_: self.destroy())
        name_entry.focus_set()
        self.grab_set()
        self.wait_window(self)

    @staticmethod
    def describe(info: dict | None, filename: str) -> str:
        """'Current: mapping.json — 5.4 KB, modified 22/10/2026 21:40' (or 'missing')."""
        if not info or not info.get("exists"):
            return f"Current: {filename} — missing"
        size = info["size"]
        size_text = f"{size / 1024:.1f} KB" if size >= 1024 else f"{size} B"
        return f"Current: {filename} — {size_text}, modified {info['modified']:%d/%m/%Y %H:%M}"

    def _file_row(self, parent, row, label, var, filename, filetypes):
        ttk.Label(parent, text=label, width=16, anchor="w").grid(row=row, column=0, sticky="w", pady=(3, 0))
        ttk.Entry(parent, textvariable=var, width=34).grid(row=row, column=1, sticky="we", pady=(3, 0))
        ttk.Button(parent, text="Browse…", width=9,
                   command=lambda: self._browse(var, label, filetypes)).grid(row=row, column=2, padx=(6, 0))
        if self.update_mode:
            ttk.Label(parent, text=self.describe(self._file_info.get(filename), filename),
                      foreground="#888888", font=("Arial", 8)).grid(
                row=row + 1, column=1, columnspan=2, sticky="w")

    def _browse(self, var, title, filetypes):
        path = filedialog.askopenfilename(parent=self, title=title, filetypes=filetypes)
        if path:
            var.set(path)

    def _on_ok(self):
        name = self.name_var.get().strip()
        if not name:
            messagebox.showerror(self.title(), "Please enter a profile name.", parent=self)
            return
        self.result = {
            "name": name,
            "mapping": self.mapping_var.get().strip() or None,
            "template": self.template_var.get().strip() or None,
            "data": self.data_var.get().strip() or None,
            "source_slug": self._existing.get(self.source_var.get()),
        }
        self.destroy()


class NewProfileDialog(ProfileDialog):
    def __init__(self, parent, existing):
        super().__init__(parent, existing)


class UpdateProfileDialog(ProfileDialog):
    def __init__(self, parent, current_name: str, file_info: dict):
        super().__init__(parent, [], current_name=current_name, file_info=file_info)
