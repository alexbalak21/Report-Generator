import re
from app.core.report_generator import ReportGenerator
from app.core.excel_reader import ExcelReader
from app.core.word_processor import WordProcessor
from app.core.mapping_loader import MappingLoader
from app.core import processors

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
report_prefix = cfg.get('report_prefix', '')
row_data = {k: rg._normalize_field_value(v, date_format) for k, v in raw_row.items()}
operations = rg._build_operations(excel, report_prefix, ROW_NUMBER)
mapping = rg.mapping_loader.load()
excel_columns = {rg._normalize_header(c) for c in excel.get_columns()}
word = WordProcessor(TEMPLATE_PATH)
available_placeholders = set(word.extract_placeholders())
filled = {}
computed_values = {'report_prefix': report_prefix}

# Pass 1
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

print('PASS1 filled includes temperature_reception?', '{{temperature_reception}}' in filled)
print('PASS1 filled includes numéro_rapport?', '{{numéro_rapport}}' in filled)
print('PASS1 filled includes date_rapport?', '{{date_rapport}}' in filled)
print('PASS1 filled', {k:v for k,v in filled.items() if k in ['{{temperature_reception}}','{{date_rapport}}','{{numéro_rapport}}']})

# Pass 2
max_passes = len(mapping) + 1
for _ in range(max_passes):
    resolved_any = False
    for key, rule in mapping.items():
        if not isinstance(rule, dict):
            continue
        if key == 'file_name':
            continue
        if key in computed_values:
            continue
        enriched = {**row_data, **computed_values}
        try:
            value = rg.compute_value(rule, enriched, operations)
        except Exception:
            value = None
        if value is None:
            continue
        computed_values[key] = str(value)
        resolved_any = True
        placeholder_full = f'{{{{{key}}}}}'
        if key in available_placeholders:
            filled[placeholder_full] = str(value)
    if not resolved_any:
        break

print('PASS2 computed_values keys', [k for k in computed_values if k in ['numero_rapport', 'date rapport']])
print('PASS2 filled includes numéro_rapport?', '{{numéro_rapport}}' in filled)
print('PASS2 filled mapping for keys with numer accents', {k:v for k,v in filled.items() if 'num' in k or 'date' in k or 'temperature' in k})

# Alias injection
num_val = computed_values.get('numero_rapport', '')
if num_val:
    filled['{{numéro_rapport}}'] = num_val
print('after alias injection', '{{numéro_rapport}}' in filled, filled.get('{{numéro_rapport}}'))

print('final filled keys count', len(filled))
print('final filled keys sample', list(filled.keys())[:20])

# Try replacing in a specific paragraph from template
p = None
for para in word._all_paragraphs():
    if '{{temperature_reception}}' in para.text or '{{numéro_rapport}}' in para.text:
        p = para
        break
if p is not None:
    print('paragraph before', p.text)
    wp = WordProcessor(TEMPLATE_PATH)
    wp._replace_in_paragraph(p, filled)
    print('paragraph after', p.text)
