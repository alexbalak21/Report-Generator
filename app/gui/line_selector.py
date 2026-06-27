"""
Reusable [ − ] [ input ] [ + ] widget for selecting an Excel row number.
"""
import tkinter as tk
from tkinter import ttk

MIN_ROW = 2  # Row 1 is always the header


class LineSelector(ttk.Frame):
    def __init__(self, parent, variable: tk.IntVar, **kwargs):
        super().__init__(parent, **kwargs)
        self._var = variable
        self._build()

    def _build(self):
        ttk.Label(self, text="Line (row ≥ 2):", width=18, anchor="w").pack(side="left")

        ttk.Button(self, text="−", width=3,
                   command=self._decrement).pack(side="left", padx=(0, 4))

        vcmd = (self.register(self._validate), "%P")
        ttk.Entry(
            self, textvariable=self._var,
            width=6, justify="center",
            validate="key", validatecommand=vcmd,
        ).pack(side="left")

        ttk.Button(self, text="+", width=3,
                   command=self._increment).pack(side="left", padx=(4, 0))

    @staticmethod
    def _validate(value: str) -> bool:
        return value == "" or (value.isdigit() and int(value) >= 1)

    def _increment(self):
        self._var.set(self._var.get() + 1)

    def _decrement(self):
        current = self._var.get()
        if current > MIN_ROW:
            self._var.set(current - 1)
