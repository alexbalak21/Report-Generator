import datetime
import unittest
import tempfile
import os
try:
    from openpyxl import Workbook
    from docx import Document
except Exception:
    Workbook = None
    Document = None
try:
    from app.core.report_generator import ReportGenerator
    from app.core.mapping_loader import MappingLoader
except Exception:
    ReportGenerator = None
    MappingLoader = None
import json


class TestReportGenerator(unittest.TestCase):
    def setUp(self):
        if Workbook is None or Document is None or ReportGenerator is None:
            self.skipTest("openpyxl, python-docx, or ReportGenerator dependencies not installed")
        self.temp_dir = tempfile.mkdtemp()
        self.xlsx_path = os.path.join(self.temp_dir, "data.xlsx")
        self.template_path = os.path.join(self.temp_dir, "template.docx")
        self.mapping_path = os.path.join(self.temp_dir, "mapping.json")
        self.output_path = os.path.join(self.temp_dir, "output.docx")

        # Create Excel workbook with one data row.
        # K-value is stored as a formula cell so the reader must resolve its computed value.
        wb = Workbook()
        ws = wb.active
        ws.append(["date rapport", "temperature reception", "imp", "k value", "date emabalage", "dlc", "numero echantillon"])
        ws.append([
            datetime.date(2026, 6, 21),
            "4°C",
            0.738026819923372,
            None,
            datetime.date(2026, 5, 31),
            datetime.date(2026, 6, 14),
            1,
        ])
        ws["D2"] = "=1-C2"
        wb.save(self.xlsx_path)

        # Create Word template with placeholders
        doc = Document()
        doc.add_paragraph("Date de rapport : {{date_rapport}}")
        doc.add_paragraph("Température : {{temperature_reception}}")
        doc.add_paragraph("IMP : {{imp}}")
        doc.add_paragraph("K-value : {{k_value}}")
        doc.add_paragraph("Emballage : {{date_emabalage}}")
        doc.add_paragraph("DLC : {{dlc}}")
        doc.add_paragraph("Report : {{numéro_rapport}}")
        doc.save(self.template_path)

        mapping = {
            "config": {
                "data_file": self.xlsx_path,
                "template_file": self.template_path,
                "output_dir": self.temp_dir,
                "date_format": "%d/%m/%Y"
            },
            "date rapport": {"column": "date rapport", "placeholder": "{{date_rapport}}"},
            "temperature reception": {"column": "temperature reception", "placeholder": "{{temperature_reception}}"},
            "imp": {
                "column": "imp",
                "placeholder": "{{imp}}",
                "operations": [
                    {"type": "formula"},
                    {"type": "multiply", "value": 100},
                    {"type": "round", "decimals": 0},
                    {"type": "suffix", "value": "%"}
                ]
            },
            "k value": {
                "column": "k value",
                "placeholder": "{{k_value}}",
                "operations": [
                    {"type": "formula"},
                    {"type": "multiply", "value": 100},
                    {"type": "round", "decimals": 0},
                    {"type": "suffix", "value": "%"}
                ]
            },
            "date emabalage": {
                "column": "date emabalage",
                "placeholder": "{{date_emabalage}}",
                "operations": [{"type": "date_format", "format": "%d/%m/%Y"}]
            },
            "dlc": {
                "column": "dlc",
                "placeholder": "{{dlc}}",
                "operations": [{"type": "date_format", "format": "%d/%m/%Y"}]
            },
            "numero echantillon": {
                "column": "numero echantillon",
                "placeholder": "{{numero_echantillon}}"
            },
            "numero_rapport": {
                "operation": "excel_day_counter",
                "date_column": "date rapport",
                "sample_column": "numero echantillon",
                "date_format": "%d/%m/%Y"
            }
        }
        with open(self.mapping_path, "w", encoding="utf-8") as f:
            json.dump(mapping, f, indent=2, ensure_ascii=False)

    def tearDown(self):
        import shutil

        if os.path.isdir(self.temp_dir):
            shutil.rmtree(self.temp_dir, ignore_errors=True)

    def test_report_generation_fills_placeholders(self):
        rg = ReportGenerator(self.xlsx_path, self.template_path, self.mapping_path)
        output = rg.generate(2, self.output_path)
        self.assertEqual(output, self.output_path)

        doc = Document(output)
        text = "\n".join(p.text for p in doc.paragraphs)
        self.assertNotIn("{{date_rapport}}", text)
        self.assertNotIn("{{temperature_reception}}", text)
        self.assertNotIn("{{imp}}", text)
        self.assertNotIn("{{k_value}}", text)
        self.assertNotIn("{{date_emabalage}}", text)
        self.assertNotIn("{{dlc}}", text)
        self.assertNotIn("{{numéro_rapport}}", text)

        self.assertIn("21/06/2026", text)
        self.assertIn("4°C", text)
        self.assertIn("74%", text)
        self.assertIn("26%", text)
        self.assertIn("31/05/2026", text)
        self.assertIn("14/06/2026", text)
        self.assertIn("260621-1", text)

    def test_report_generation_with_xlsx_mapping(self):
        if MappingLoader is None:
            self.skipTest("MappingLoader unavailable")

        mapping_xlsx_path = os.path.join(self.temp_dir, "mapping.xlsx")
        wb = Workbook()
        ws = wb.active
        ws.title = "mappings"
        ws.append(["Spreadsheet Column", "Placeholder", "Operation", "Type", "Notes"])
        ws.append(["date rapport", "{{date_rapport}}", "", "date (input)", ""])
        ws.append(["temperature reception", "{{temperature_reception}}", "", "text (input)", ""])
        ws.append(["date emabalage", "{{date_emabalage}}", "date_format %d/%m/%Y", "date", ""])
        ws.append(["dlc", "{{dlc}}", "date_format %d/%m/%Y", "date", ""])
        ws.append(["numero echantillon", "{{numero_echantillon}}", "", "text/number (input)", ""])
        ws.append(["(computed) numero_rapport", "—", "report_number (date_column: date rapport; sample_column: numero echantillon)", "computed", ""])
        ws.append(["(computed) file_name", "—", "format(\"{name} {numero_rapport}.docx\")", "computed", ""])
        wb.save(mapping_xlsx_path)

        loader = MappingLoader(mapping_xlsx_path)
        loader.update_file_name_field(name="TEST")
        loaded = loader.load()
        self.assertEqual(loaded["file_name"]["name"], "TEST")
        self.assertEqual(loaded["file_name"]["operation"], "format")

        output = ReportGenerator(self.xlsx_path, self.template_path, mapping_xlsx_path).generate(2, self.output_path)
        self.assertEqual(output, self.output_path)

        doc = Document(output)
        text = "\n".join(p.text for p in doc.paragraphs)
        self.assertIn("21/06/2026", text)
        self.assertIn("4°C", text)
        self.assertIn("260621-1", text)


if __name__ == "__main__":
    unittest.main()
