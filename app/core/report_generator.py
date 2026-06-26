import json
import re
import datetime
from .excel_reader import ExcelReader
from .word_processor import WordProcessor


class ReportGenerator:
    def __init__(self, excel_path, template_path, mapping_path):
        self.excel_path = excel_path
        self.template_path = template_path
        self.mapping_path = mapping_path

    @staticmethod
    def _normalize_header(value):
        if value is None:
            return ""
        return str(value).strip()

    def load_mapping(self):
        """Load mapping.json from /mappings."""
        with open(self.mapping_path, "r", encoding="utf-8") as f:
            return json.load(f)

    # ---------------------------------------------------------
    # COMPUTED FIELD ENGINE
    # ---------------------------------------------------------
    def compute_value(self, rule, row_data):
        operation = rule.get("operation")

        # ---- TODAY ----
        if operation == "today":
            fmt = rule.get("format", "%Y-%m-%d")
            return datetime.date.today().strftime(fmt)

        # ---- UPPERCASE ----
        if operation == "uppercase":
            input_field = rule.get("input")
            if input_field in row_data:
                return str(row_data[input_field]).upper()

        # ---- LOWERCASE ----
        if operation == "lowercase":
            input_field = rule.get("input")
            if input_field in row_data:
                return str(row_data[input_field]).lower()

        # ---- FORMAT STRING ----
        if operation == "format":
            fmt = rule.get("format", "")
            try:
                return fmt.format(**row_data)
            except Exception:
                return ""

        # ---- CONCAT ----
        if operation == "concat":
            parts = rule.get("parts", [])
            return "".join(str(row_data.get(p, "")) for p in parts)

        # ---- DEFAULT ----
        return None

    # ---------------------------------------------------------
    # MAIN GENERATION LOGIC
    # ---------------------------------------------------------
    def generate(self, row_number, output_path):
        excel = ExcelReader(self.excel_path)
        excel.load()
        row_data = excel.get_row_as_dict(row_number)

        mapping = self.load_mapping()
        excel_columns = {self._normalize_header(column) for column in excel.get_columns()}

        word = WordProcessor(self.template_path)
        available_placeholders = set(word.extract_placeholders())

        filled = {}

        for key, rule in mapping.items():

            # -------------------------------------------------
            # CASE 1: SIMPLE PLACEHOLDER (string)
            # -------------------------------------------------
            if isinstance(rule, str):
                normalized_key = self._normalize_header(key)

                # skip if Excel column doesn't exist
                if normalized_key not in excel_columns:
                    continue

                # skip if row doesn't contain the column
                if normalized_key not in row_data:
                    continue

                # extract placeholder name inside {{ }}
                match = re.search(r"\{\{(.*?)\}\}", rule)
                if not match:
                    continue

                placeholder_name = match.group(1).strip()

                # skip if placeholder not in Word template
                if placeholder_name not in available_placeholders:
                    continue

                filled[rule] = row_data[normalized_key]
                continue

            # -------------------------------------------------
            # CASE 2: COMPUTED FIELD (dict)
            # -------------------------------------------------
            if isinstance(rule, dict):
                value = self.compute_value(rule, row_data)

                if value is None:
                    continue

                # placeholder name = key
                placeholder_name = key

                # Word expects {{placeholder}}
                placeholder_full = f"{{{{{placeholder_name}}}}}"

                if placeholder_name not in available_placeholders:
                    continue

                filled[placeholder_full] = value
                continue

        # ---------------------------------------------------------
        # APPLY TO WORD TEMPLATE
        # ---------------------------------------------------------
        word.fill_placeholders(filled, output_path)

        return output_path
