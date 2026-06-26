import tkinter as tk
from tkinter import filedialog


def select_excel_file() -> str:
    """Open a file dialog to select an Excel .xlsx file."""
    return filedialog.askopenfilename(
        title="Select Excel Data File",
        filetypes=[("Excel Files", "*.xlsx")],
    )


def select_docx_template() -> str:
    """Open a file dialog to select a Word .docx template."""
    return filedialog.askopenfilename(
        title="Select Word Template File",
        filetypes=[("Word Documents", "*.docx")],
    )


def select_mapping_file() -> str:
    """Open a file dialog to select a JSON mapping/configuration file."""
    return filedialog.askopenfilename(
        title="Select Mapping Configuration File",
        filetypes=[("JSON Files", "*.json")],
    )
