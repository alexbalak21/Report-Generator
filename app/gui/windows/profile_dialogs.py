"""
Dialogs used by the profile selector (kept apart from main_window.py).
"""
import tkinter as tk
from tkinter import ttk, filedialog, messagebox


class NewProfileDialog(tk.Toplevel):
    """
    Ask for a profile name and (optionally) its three files.
    After the dialog closes, `self.result` is None (cancelled) or a dict:
        {"name", "mapping", "template", "data", "source_slug"}
    Empty file fields stay empty; they can be added later with the ⋯ menu.
    """

    _NO_SOURCE = "(empty profile)"

    def __init__(self, parent, existing: list[tuple[str, str]]):
        """existing — [(slug, display name), …] of profiles to copy from."""
        super().__init__(parent)
        self.title("New profile")
        self.resizable(False, False)
        self.transient(parent)
        self.result = None
        self._existing = dict((name, slug) for slug, name in existing)

        self.name_var     = tk.StringVar()
        self.source_var   = tk.StringVar(value=self._NO_SOURCE)
        self.mapping_var  = tk.StringVar()
        self.template_var = tk.StringVar()
        self.data_var     = tk.StringVar()

        body = ttk.Frame(self, padding=14)
        body.pack(fill="both", expand=True)

        ttk.Label(body, text="Profile name", width=16, anchor="w").grid(row=0, column=0, sticky="w", pady=3)
        name_entry = ttk.Entry(body, textvariable=self.name_var, width=44)
        name_entry.grid(row=0, column=1, columnspan=2, sticky="we", pady=3)

        ttk.Label(body, text="Start from", width=16, anchor="w").grid(row=1, column=0, sticky="w", pady=3)
        ttk.Combobox(
            body, textvariable=self.source_var, state="readonly", width=41,
            values=[self._NO_SOURCE, *self._existing.keys()],
        ).grid(row=1, column=1, columnspan=2, sticky="we", pady=3)

        ttk.Separator(body).grid(row=2, column=0, columnspan=3, sticky="we", pady=8)
        ttk.Label(body, text="Files (optional — copied into the profile)",
                  foreground="#666666").grid(row=3, column=0, columnspan=3, sticky="w")

        self._file_row(body, 4, "Mapping (.json)", self.mapping_var,
                       [("Mapping files", "*.json *.xlsx")])
        self._file_row(body, 5, "Template (.docx)", self.template_var,
                       [("Word documents", "*.docx")])
        self._file_row(body, 6, "Data (.xlsx)", self.data_var,
                       [("Excel files", "*.xlsx *.xlsm")])

        buttons = ttk.Frame(body)
        buttons.grid(row=7, column=0, columnspan=3, sticky="e", pady=(12, 0))
        ttk.Button(buttons, text="Cancel", command=self.destroy).pack(side="right")
        ttk.Button(buttons, text="Create", command=self._on_create).pack(side="right", padx=(0, 8))

        self.bind("<Return>", lambda *_: self._on_create())
        self.bind("<Escape>", lambda *_: self.destroy())
        name_entry.focus_set()
        self.grab_set()
        self.wait_window(self)

    def _file_row(self, parent, row, label, var, filetypes):
        ttk.Label(parent, text=label, width=16, anchor="w").grid(row=row, column=0, sticky="w", pady=3)
        ttk.Entry(parent, textvariable=var, width=34).grid(row=row, column=1, sticky="we", pady=3)
        ttk.Button(parent, text="Browse…", width=9,
                   command=lambda: self._browse(var, label, filetypes)).grid(row=row, column=2, padx=(6, 0))

    def _browse(self, var, title, filetypes):
        path = filedialog.askopenfilename(parent=self, title=title, filetypes=filetypes)
        if path:
            var.set(path)

    def _on_create(self):
        name = self.name_var.get().strip()
        if not name:
            messagebox.showerror("New profile", "Please enter a profile name.", parent=self)
            return
        self.result = {
            "name": name,
            "mapping": self.mapping_var.get().strip() or None,
            "template": self.template_var.get().strip() or None,
            "data": self.data_var.get().strip() or None,
            "source_slug": self._existing.get(self.source_var.get()),
        }
        self.destroy()
