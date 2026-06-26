import json
from .excel_reader import ExcelReader
from .word_processor import WordProcessor

class ReportGenerator:
    def __init__(self, excel_path, template_path, mapping_path):
        self.excel_path = excel_path
        self.template_path = template_path
        self.mapping_path = mapping_path

    def load_mapping(self):
        """Load mapping.json from /mappings."""
        with open(self.mapping_path, "r", encoding="utf-8") as f:
            return json.load(f)

    def generate(self, row_number, output_path):
        """
        Generate a report using:
        - Excel row
        - Word template
        - mapping.json
        """
        # Load Excel
        excel = ExcelReader(self.excel_path)
        row_data = excel.get_row_as_dict(row_number)

        # Load mapping.json
        mapping = self.load_mapping()

        # Build final placeholder → value dict
        filled = {
            placeholder: row_data[column]
            for column, placeholder in mapping.items()
        }

        # Process Word template
        word = WordProcessor(self.template_path)
        word.fill_placeholders(filled, output_path)

        return output_path
