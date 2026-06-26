import tempfile
import unittest
from pathlib import Path

import openpyxl
from docx import Document

from app.core.report_generator import ReportGenerator


class ReportGeneratorTests(unittest.TestCase):
    def test_generate_raises_error_for_mismatched_mapping_and_columns(self):
        with tempfile.TemporaryDirectory() as tmp_dir:
            excel_path = Path(tmp_dir) / "data.xlsx"
            template_path = Path(tmp_dir) / "template.docx"
            mapping_path = Path(tmp_dir) / "mapping.json"
            output_path = Path(tmp_dir) / "output.docx"

            workbook = openpyxl.Workbook()
            sheet = workbook.active
            sheet.append(["date", "species"])
            workbook.save(excel_path)

            template_doc = Document()
            template_doc.add_paragraph("{{date}}")
            template_doc.save(template_path)

            mapping_path.write_text(
                '{"date": "{{date}}", "missing_column": "{{missing}}"}',
                encoding="utf-8",
            )

            generator = ReportGenerator(str(excel_path), str(template_path), str(mapping_path))

            with self.assertRaisesRegex(ValueError, "missing_column|species"):
                generator.generate(1, str(output_path))


if __name__ == "__main__":
    unittest.main()
