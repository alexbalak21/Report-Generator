import re
import copy
try:
    from docx import Document
    from docx.oxml.ns import qn
except Exception:
    Document = None
    qn = None

PLACEHOLDER_RE = re.compile(r"\{\{.*?\}\}")


class WordProcessor:
    def __init__(self, filepath):
        if Document is None:
            raise ImportError("python-docx is required for WordProcessor")
        self.filepath = filepath
        self.document = Document(filepath)

    # ------------------------------------------------------------------
    # Iterate every paragraph in the document including tables
    # ------------------------------------------------------------------

    def _all_paragraphs(self):
        yield from self.document.paragraphs
        for table in self.document.tables:
            yield from self._paragraphs_in_table(table)

    def _paragraphs_in_table(self, table):
        for row in table.rows:
            for cell in row.cells:
                yield from cell.paragraphs
                for nested in cell.tables:
                    yield from self._paragraphs_in_table(nested)

    # ------------------------------------------------------------------
    # Public API
    # ------------------------------------------------------------------

    def extract_placeholders(self) -> list:
        found = []
        for para in self._all_paragraphs():
            found.extend(re.findall(r"\{\{(.*?)\}\}", para.text))
        return found

    def fill_placeholders(self, mapping: dict, output_path: str) -> None:
        for para in self._all_paragraphs():
            self._replace_in_paragraph(para, mapping)
            if "{{" in para.text:
                # Second pass may be needed for placeholders that were split across
                # runs in a way the first merge pass did not fully consolidate.
                self._replace_in_paragraph(para, mapping)
        self.document.save(output_path)

    # ------------------------------------------------------------------
    # Core replacement — preserves per-run formatting
    # ------------------------------------------------------------------

    def _replace_in_paragraph(self, paragraph, mapping: dict) -> None:
        """
        Replace placeholders while preserving the formatting of the run
        that contains (or starts) each placeholder.

        Algorithm:
        1. Merge runs that together form a split placeholder into a single run
           (inheriting the first fragment's formatting).
        2. For each run, do simple string replacement — the run keeps its own
           bold/color/font because we never touch runs that don't contain a tag.
        """
        if not paragraph.runs:
            return

        full_text = "".join(r.text for r in paragraph.runs)
        if "{{" not in full_text:
            return

        # ── Step 1: consolidate split-run placeholders ─────────────────
        self._merge_split_placeholders(paragraph)

        # ── Step 2: replace in each run individually ───────────────────
        for run in paragraph.runs:
            if "{{" not in run.text:
                continue
            for placeholder, value in mapping.items():
                if placeholder in run.text:
                    run.text = run.text.replace(placeholder, str(value))

    def _merge_split_placeholders(self, paragraph) -> None:
        """
        Scan runs left-to-right.  When an opening '{{' has no matching '}}'
        in the same run, keep absorbing subsequent runs (copying their text
        into the first fragment's run, inheriting its formatting) until the
        closing '}}' is found.
        """
        runs = paragraph.runs
        i = 0
        while i < len(runs):
            text = runs[i].text
            open_idx = text.find("{{")
            if open_idx == -1:
                i += 1
                continue

            # Check if the closing braces are already in this run
            if "}}" in text[open_idx:]:
                i += 1
                continue

            # Need to absorb subsequent runs until we find '}}'
            j = i + 1
            accumulated = text
            while j < len(runs):
                accumulated += runs[j].text
                if "}}" in accumulated:
                    break
                j += 1

            if "}}" not in accumulated:
                # Malformed / incomplete placeholder — leave as-is
                i += 1
                continue

            # Write accumulated text into run[i], clear runs i+1..j
            runs[i].text = accumulated
            for k in range(i + 1, j + 1):
                runs[k].text = ""

            # Don't advance i — the merged run might contain multiple placeholders
            # but the outer loop will handle them in Step 2.
            i += 1
