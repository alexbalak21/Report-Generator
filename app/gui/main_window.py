import tkinter as tk
from tkinter import ttk

from .file_dialogs import select_excel_file, select_docx_template


class MainWindow(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title("Report Generator")
        self.geometry("500x250")
        self.resizable(False, False)

        # Store selected paths
        self.excel_path = tk.StringVar()
        self.template_path = tk.StringVar()

        # UI
        self.create_widgets()

    def create_widgets(self):
        # Title
        title = ttk.Label(self, text="Simple Report Generator", font=("Arial", 16))
        title.pack(pady=10)

        # Select Excel button
        btn_excel = ttk.Button(self, text="Select data .xlsx", command=self.on_select_excel)
        btn_excel.pack(pady=5)

        lbl_excel = ttk.Label(self, textvariable=self.excel_path, foreground="blue")
        lbl_excel.pack()

        # Select Template button
        btn_template = ttk.Button(self, text="Select template .docx", command=self.on_select_template)
        btn_template.pack(pady=5)

        lbl_template = ttk.Label(self, textvariable=self.template_path, foreground="blue")
        lbl_template.pack()

    def on_select_excel(self):
        path = select_excel_file()
        if path:
            self.excel_path.set(path)

    def on_select_template(self):
        path = select_docx_template()
        if path:
            self.template_path.set(path)


if __name__ == "__main__":
    app = MainWindow()
    app.mainloop()
