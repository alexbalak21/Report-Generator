import os
import re
from .excel_reader import ExcelReader
from .word_processor import WordProcessor
from .mapping_loader import MappingLoader
from .state_manager import ReportStateManager
from . import processors


class ReportGenerator:
    def __init__(self, excel_path, template_path, mapping_path):
        self.excel_path = excel_path
        self.template_path = template_path
        self.mapping_path = mapping_path

        self.mapping_loader = MappingLoader(mapping_path)
        self.state_manager = ReportStateManager(
            os.path.join(os.path.dirname(mapping_path), "report_state.json")
        )

        # Map operation names → functions
        self.operations = {
            "today": processors.op_today,
            "uppercase": processors.op_uppercase,
            "lowercase": processors.op_lowercase,
            "format": processors.op_format,
            "concat": processors.op_concat,
            "report_number": lambda rule, row: self.state_manager.generate_report_number()
        }

    @staticmethod
    def _normalize_header(value):
        if value is None:
            return ""
        return str(value).strip()

    def compute_value(self, rule, row_data):
        operation = rule.get("operation")
        func = self.operations.get(operation)

        if func:
            return func(rule, row_data)

        return None

    def generate(self, row_number, output_path):
        excel = ExcelReader(self.excel_path)
        excel.load()
        row_data = excel.get_row_as_dict(row_number)

        mapping = self.mapping_loader.load()
        excel_columns = {self._normalize_header(c) for c in excel.get_columns()}

        word = WordProcessor(self.template_path)
        available_placeholders = set(word.extract_placeholders())

        filled = {}

        for key, rule in mapping.items():

            # SIMPLE PLACEHOLDER
            if isinstance(rule, str):
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

                filled[rule] = row_data[normalized_key]
                continue

            # COMPUTED FIELD
            if isinstance(rule, dict):
                value = self.compute_value(rule, row_data)
                if value is None:
                    continue

                placeholder_full = f"{{{{{key}}}}}"

                if key not in available_placeholders:
                    continue

                filled[placeholder_full] = value

        word.fill_placeholders(filled, output_path)
        return output_path
