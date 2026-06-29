from app.core.report_generator import ReportGenerator
from app.core.word_processor import WordProcessor

orig_fill = WordProcessor.fill_placeholders

def debug_fill(self, mapping, output_path):
    print('DEBUG fill mapping keys', sorted(mapping.keys()))
    print('DEBUG values', {k: mapping[k] for k in mapping if k in ['{{temperature_reception}}','{{numéro_rapport}}','{{date_rapport}}']})
    return orig_fill(self, mapping, output_path)

WordProcessor.fill_placeholders = debug_fill

rg = ReportGenerator('data/final_rapport_data.xlsx', 'data/rapport_template.docx', 'new_rapport_mapping.json')
print('GENERATE', rg.generate(2, 'out/test_output_debug.docx'))
