from app.core.report_generator import ReportGenerator
from app.core.excel_reader import ExcelReader
from app.core.word_processor import WordProcessor
from app.core.mapping_loader import MappingLoader
from app.core import processors
import re

EXCEL_PATH = 'data/final_rapport_data.xlsx'
TEMPLATE_PATH = 'data/rapport_template.docx'
MAPPING_PATH = 'new_rapport_mapping.json'
ROW_NUMBER = 2

rg = ReportGenerator(EXCEL_PATH, TEMPLATE_PATH, MAPPING_PATH)
excel = ExcelReader(EXCEL_PATH)
excel.load()
raw_row = excel.get_row_as_dict(ROW_NUMBER)

cfg = rg.mapping_loader.load_config()
date_format = cfg.get('date_format', '%d/%m/%Y')
row_data = {k: rg._normalize_field_value(v, date_format) for k, v in raw_row.items()}

operations = rg._build_operations(excel, cfg.get('report_prefix', ''), ROW_NUMBER)
mapping = rg.mapping_loader.load()
excel_columns = {rg._normalize_header(c) for c in excel.get_columns()}
word = WordProcessor(TEMPLATE_PATH)
available_placeholders = set(word.extract_placeholders())
filled = {}
computed_values = {'report_prefix': cfg.get('report_prefix', '')}

for key, rule in mapping.items():
    if isinstance(rule, str):
        normalized_key = rg._normalize_header(key)
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
        continue
    if isinstance(rule, dict) and 'column' in rule and 'placeholder' in rule:
        column_name = rg._normalize_header(rule.get('column', ''))
        if column_name not in excel_columns:
            continue
        value = raw_row.get(column_name)
        try:
            value_processed = processors.apply_operations(
                value, rule.get('operations'), {**row_data, **computed_values}, excel
            )
        except Exception:
            value_processed = value
        placeholder = rule.get('placeholder')
        match = re.search(r"\{\{(.*?)\}\}", placeholder) if isinstance(placeholder, str) else None
        if not match:
            continue
        placeholder_name = match.group(1).strip()
        if placeholder_name not in available_placeholders:
            continue
        filled[placeholder] = value_processed
        computed_values[key] = value_processed

print('available_placeholders', sorted(available_placeholders))
print('filled keys', sorted(filled.keys()))
print('filled values', filled)
print('computed_values keys', sorted(computed_values.keys()))
print('computed_values', computed_values)
print('raw_row date rapport', raw_row.get('date rapport'), type(raw_row.get('date rapport')))
print('row_data date rapport', row_data.get('date rapport'), type(row_data.get('date rapport')))

for key in ['date rapport', 'temperature reception', 'numero_rapport']:
    rule = mapping.get(key)
    print('\nrule', key, rule)
