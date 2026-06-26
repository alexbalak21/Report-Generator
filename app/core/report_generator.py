import json
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

    def validate_mapping(self, excel, mapping):
        """Ensure every mapping column exists in the Excel file and vice versa."""
        excel_columns = {self._normalize_header(column) for column in excel.get_columns()}
        mapping_columns = {self._normalize_header(column) for column in mapping.keys()}

        missing_from_excel = sorted(mapping_columns - excel_columns)
        missing_from_mapping = sorted(excel_columns - mapping_columns)

        if missing_from_excel or missing_from_mapping:
            details = []
            if missing_from_excel:
                details.append(f"columns in mapping not found in Excel: {', '.join(missing_from_excel)}")
            if missing_from_mapping:
                details.append(f"columns in Excel not found in mapping: {', '.join(missing_from_mapping)}")
            raise ValueError("Mapping mismatch: " + "; ".join(details))

    def generate(self, row_number, output_path):
        """
        Generate a report using:
        - Excel row
        - Word template
        - mapping.json
        """
        excel = ExcelReader(self.excel_path)
        excel.load()
        row_data = excel.get_row_as_dict(row_number)

        mapping = self.load_mapping()
        self.validate_mapping(excel, mapping)

        filled = {
            placeholder: row_data[self._normalize_header(column)]
            for column, placeholder in mapping.items()
        }

        word = WordProcessor(self.template_path)
        word.fill_placeholders(filled, output_path)

        return output_path
