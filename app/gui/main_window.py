import os
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog

from app.core.report_generator import ReportGenerator
from app.gui.file_dialogs import select_excel_file, select_docx_template


class MainWindow(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title("Report Generator")
        self.geometry("550x350")
        self.resizable(False, False)

        # Store selected paths
        self.excel_path = tk.StringVar()
        self.template_path = tk.StringVar()
        self.selected_row = tk.StringVar()

        self.create_widgets()

    def create_widgets(self):
        title = ttk.Label(self, text="Simple Report Generator", font=("Arial", 16))
        title.pack(pady=10)

        # Excel selection
        ttk.Button(self, text="Select data .xlsx", command=self.on_select_excel).pack(pady=5)
        ttk.Label(self, textvariable=self.excel_path, foreground="blue").pack()

        # Template selection
        ttk.Button(self, text="Select template .docx", command=self.on_select_template).pack(pady=5)
        ttk.Label(self, textvariable=self.template_path, foreground="blue").pack()

        # Row selection
        ttk.Button(self, text="Select line (row)", command=self.on_select_row).pack(pady=10)
        ttk.Label(self, textvariable=self.selected_row, foreground="green").pack()

        # Generate report
        ttk.Button(self, text="Generate report", command=self.on_generate_report).pack(pady=20)

    # -----------------------------
    # Button callbacks
    # -----------------------------

    def on_select_excel(self):
        path = select_excel_file()
        if path:
            self.excel_path.set(path)

    def on_select_template(self):
        path = select_docx_template()
        if path:
            self.template_path.set(path)

    def on_select_row(self):
        if not self.excel_path.get():
            messagebox.showerror("Error", "Select an Excel file first")
            return

        # Ask user for row number
        row = simpledialog.askinteger("Select row", "Enter Excel row number (starting at 2):")

        if row and row >= 2:
            self.selected_row.set(f"Selected row: {row}")
        else:
            messagebox.showerror("Error", "Invalid row number")

    def on_generate_report(self):
        if not self.excel_path.get() or not self.template_path.get():
            messagebox.showerror("Error", "Select Excel and Template first")
            return

        if not self.selected_row.get():
            messagebox.showerror("Error", "Select a row first")
            return

        row_number = int(self.selected_row.get().split(":")[1].strip())

        # Output file
        output_path = os.path.join(os.getcwd(), f"report_row_{row_number}.docx")

        mapping_path = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "mappings", "data.json"))

        generator = ReportGenerator(
            excel_path=self.excel_path.get(),
            template_path=self.template_path.get(),
            mapping_path=mapping_path
        )

        generator.generate(row_number=row_number, output_path=output_path)

        messagebox.showinfo("Success", f"Report generated:\n{output_path}")


if __name__ == "__main__":
    app = MainWindow()
    app.mainloop()
