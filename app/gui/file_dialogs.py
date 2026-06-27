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


def select_output_file(initial_dir: str = "", initial_file: str = "") -> str:
    """
    Open a Save As dialog so the user can choose where to save the report.
    Returns the chosen path, or empty string if cancelled.
    """
    return filedialog.asksaveasfilename(
        title="Save Report As",
        initialdir=initial_dir or None,
        initialfile=initial_file or None,
        defaultextension=".docx",
        filetypes=[("Word Documents", "*.docx")],
    )