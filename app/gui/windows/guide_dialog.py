"""
📖 Guide window: shows the bundled guide (markdown, as plain text with light styling)
and copies it — alone, or together with the active profile — for an AI assistant.
"""
import tkinter as tk
from tkinter import ttk, messagebox

from app.core.guide import load_guide, headings, build_ai_message


class GuideDialog(tk.Toplevel):
    def __init__(self, parent, context_provider=None):
        """
        context_provider — callable returning the markdown description of the active
                           profile (or None / raising if there is no profile).
        """
        super().__init__(parent)
        self.title("Guide — Report Generator")
        self.geometry("860x640")
        self.minsize(560, 360)
        self._context_provider = context_provider
        self._guide = load_guide()
        self._reset_jobs = {}

        self._build_ui()
        self._render(self._guide)
        self.bind("<Escape>", lambda *_: self.destroy())

    # ------------------------------------------------------------------

    def _build_ui(self):
        top = ttk.Frame(self, padding=(10, 8, 10, 4))
        top.pack(fill="x")

        ttk.Label(top, text="Go to:").pack(side="left")
        self._sections = headings(self._guide, level=2)
        self._jump_var = tk.StringVar()
        self._jump = ttk.Combobox(top, textvariable=self._jump_var, state="readonly", width=48,
                                  values=[title for title, _ in self._sections])
        self._jump.pack(side="left", padx=(6, 12))
        self._jump.bind("<<ComboboxSelected>>", lambda *_: self._on_jump())

        ttk.Label(top, text="Find:").pack(side="left")
        self._find_var = tk.StringVar()
        find = ttk.Entry(top, textvariable=self._find_var, width=18)
        find.pack(side="left", padx=(6, 4))
        find.bind("<Return>", lambda *_: self._find_next())
        ttk.Button(top, text="Next", width=6, command=self._find_next).pack(side="left")

        body = ttk.Frame(self, padding=(10, 0, 10, 0))
        body.pack(fill="both", expand=True)
        self._text = tk.Text(body, wrap="word", font=("Segoe UI", 10), padx=12, pady=8,
                             borderwidth=1, relief="solid", spacing1=1, spacing3=1)
        scroll = ttk.Scrollbar(body, orient="vertical", command=self._text.yview)
        self._text.configure(yscrollcommand=scroll.set)
        scroll.pack(side="right", fill="y")
        self._text.pack(side="left", fill="both", expand=True)

        self._text.tag_configure("h1", font=("Segoe UI", 17, "bold"), spacing1=6, spacing3=8)
        self._text.tag_configure("h2", font=("Segoe UI", 13, "bold"), spacing1=14, spacing3=4,
                                 foreground="#1a5fb4")
        self._text.tag_configure("h3", font=("Segoe UI", 11, "bold"), spacing1=8, spacing3=2)
        self._text.tag_configure("code", font=("Consolas", 9), background="#f3f3f3")
        self._text.tag_configure("table", font=("Consolas", 9))
        self._text.tag_configure("quote", foreground="#555555", lmargin1=14, lmargin2=14)
        self._text.tag_configure("found", background="#ffe58a")

        bottom = ttk.Frame(self, padding=(10, 8, 10, 10))
        bottom.pack(fill="x")
        self._copy_btn = ttk.Button(bottom, text="📋 Copy guide", command=self.on_copy_guide)
        self._copy_btn.pack(side="left")
        self._copy_ai_btn = ttk.Button(bottom, text="📋 Copy guide + my profile",
                                       command=self.on_copy_with_profile)
        self._copy_ai_btn.pack(side="left", padx=(8, 0))
        ttk.Label(bottom, text="Paste it in an AI chat to get help fixing your JSON, Word or Excel files.",
                  foreground="#666666").pack(side="left", padx=(12, 0))
        ttk.Button(bottom, text="Close", command=self.destroy).pack(side="right")

    # ------------------------------------------------------------------
    # Rendering: markdown shown as plain text, headings and code styled
    # ------------------------------------------------------------------

    def _render(self, text: str):
        box = self._text
        box.configure(state="normal")
        box.delete("1.0", "end")
        self._section_marks = []                    # sec_0, sec_1, … in the order of self._sections
        in_code = False
        for line in text.splitlines():
            if line.startswith("```"):
                in_code = not in_code
                continue
            if in_code:
                box.insert("end", line + "\n", "code")
            elif line.startswith("### "):
                box.insert("end", line[4:] + "\n", "h3")
            elif line.startswith("## "):
                self._section_marks.append(f"sec_{len(self._section_marks)}")
                box.mark_set(self._section_marks[-1], "end-1c")
                box.mark_gravity(self._section_marks[-1], "left")
                box.insert("end", line[3:] + "\n", "h2")
            elif line.startswith("# "):
                box.insert("end", line[2:] + "\n", "h1")
            elif line.startswith("|"):
                box.insert("end", line + "\n", "table")
            elif line.startswith(">"):
                box.insert("end", line.lstrip("> ").rstrip() + "\n", "quote")
            else:
                box.insert("end", line.replace("**", "").replace("`", "") + "\n")
        box.configure(state="disabled")

    def _on_jump(self):
        index = self._jump.current()
        if 0 <= index < len(self._section_marks):
            self._text.see("end")                   # scroll so the target lands at the top
            self._text.yview(self._section_marks[index])

    def _find_next(self):
        needle = self._find_var.get().strip()
        box = self._text
        box.tag_remove("found", "1.0", "end")
        if not needle:
            return
        pos = box.search(needle, "insert", nocase=True, stopindex="end") or \
            box.search(needle, "1.0", nocase=True, stopindex="end")     # wrap around
        if not pos:
            return
        end = f"{pos}+{len(needle)}c"
        box.tag_add("found", pos, end)
        box.mark_set("insert", end)
        box.see(pos)

    # ------------------------------------------------------------------
    # Clipboard
    # ------------------------------------------------------------------

    def _to_clipboard(self, text: str, button, original: str):
        self.clipboard_clear()
        self.clipboard_append(text)
        self.update()                               # keep the text after the window closes
        button.configure(text="✔ Copied")
        job = self._reset_jobs.get(button)
        if job:
            self.after_cancel(job)
        self._reset_jobs[button] = self.after(1800, lambda: button.configure(text=original))

    def on_copy_guide(self):
        self._to_clipboard(self._guide, self._copy_btn, "📋 Copy guide")

    def on_copy_with_profile(self):
        context = None
        if self._context_provider:
            try:
                context = self._context_provider()
            except Exception as exc:
                messagebox.showwarning(
                    "Guide", f"Your profile could not be read ({exc}).\nThe guide alone was copied.",
                    parent=self)
        self._to_clipboard(build_ai_message(self._guide, context), self._copy_ai_btn,
                           "📋 Copy guide + my profile")
