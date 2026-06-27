import os
import re
from .excel_reader import ExcelReader
from .word_processor import WordProcessor
from .mapping_loader import MappingLoader
from .state_manager import ReportStateManager
from . import processors

# Characters illegal in Windows filenames
_ILLEGAL_FILENAME_CHARS = re.compile(r'[<>:"/\\|?*\x00-\x1f]')


class ReportGenerator:
    def __init__(self, excel_path, template_path, mapping_path):
        self.excel_path = excel_path
        self.template_path = template_path
        self.mapping_path = mapping_path

        self.mapping_loader = MappingLoader(mapping_path)
        self.state_manager = ReportStateManager(
            os.path.join(os.path.dirname(mapping_path), "report_state.json")
        )

        self.operations = {
            "today":         processors.op_today,
            "uppercase":     processors.op_uppercase,
            "lowercase":     processors.op_lowercase,
            "format":        processors.op_format,
            "concat":        processors.op_concat,
            "report_number": lambda rule, row: self.state_manager.generate_report_number(),
        }

    @staticmethod
    def _normalize_header(value):
        if value is None:
            return ""
        return str(value).strip()

    @staticmethod
    def _sanitize_filename(name: str) -> str:
        """Remove characters illegal in Windows filenames and collapse spaces."""
        sanitized = _ILLEGAL_FILENAME_CHARS.sub("_", name)
        sanitized = sanitized.strip(". ")   # no leading/trailing dots or spaces
        return sanitized or "report"

    @staticmethod
    def _normalize_field_value(value) -> str:
        """
        Convert an Excel cell value to a clean string suitable for use in
        filenames and format strings. Datetime values are formatted as YYYY-MM-DD.
        """
        import datetime as dt
        if isinstance(value, (dt.datetime, dt.date)):
            return value.strftime("%Y-%m-%d")
        if value is None:
            return ""
        return str(value).strip()

    def compute_value(self, rule, row_data):
        operation = rule.get("operation")
        func = self.operations.get(operation)
        return func(rule, row_data) if func else None

    def generate(self, row_number, output_path):
        excel = ExcelReader(self.excel_path)
        excel.load()
        raw_row = excel.get_row_as_dict(row_number)

        # Normalize all values (converts datetime → "YYYY-MM-DD", None → "")
        row_data = {k: self._normalize_field_value(v) for k, v in raw_row.items()}

        mapping = self.mapping_loader.load()
        excel_columns = {self._normalize_header(c) for c in excel.get_columns()}

        word = WordProcessor(self.template_path)
        available_placeholders = set(word.extract_placeholders())

        filled = {}
        computed_values = {}

        # ── Pass 1: simple column mappings ────────────────────────────
        for key, rule in mapping.items():
            if not isinstance(rule, str):
                continue
            normalized_key = self._normalize_header(key)
            if normalized_key not in excel_columns:
                continue
            if normalized_key not in row_data:
                continue
            match = re.search(r"\{\{(.*?)\}\}", rule)
            if not match:
                continue
            placeholder_name = match.group(1).strip()
            if placeholder_name not in available_placeholders:
                continue
            value = row_data[normalized_key]
            filled[rule] = value
            computed_values[key] = value

        # ── Pass 2: computed fields (excluding file_name) ─────────────
        for key, rule in mapping.items():
            if not isinstance(rule, dict):
                continue
            if key == "file_name":
                continue  # resolved last, after report_number is known

            enriched_row = {**row_data, **computed_values}
            value = self.compute_value(rule, enriched_row)
            if value is None:
                continue

            computed_values[key] = str(value)

            placeholder_full = f"{{{{{key}}}}}"
            if key not in available_placeholders:
                continue
            filled[placeholder_full] = str(value)

        # ── Write output ──────────────────────────────────────────────
        # output_path is the explicit path chosen by the user via Save As dialog
        os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
        word.fill_placeholders(filled, output_path)
        return output_path