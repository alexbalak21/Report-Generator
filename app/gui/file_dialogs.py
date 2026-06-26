import tkinter as tk
from tkinter import filedialog

def select_excel_file():
    """Open a file dialog to select an Excel .xlsx file."""
    filepath = filedialog.askopenfilename(
        title="Select Excel Data File",
        filetypes=[("Excel Files", "*.xlsx")]
    )
    return filepath

def select_docx_template():
    """Open a file dialog to select a Word .docx template."""
    filepath = filedialog.askopenfilename(
        title="Select Word Template File",
        filetypes=[("Word Documents", "*.docx")]
    )
    return filepath
