# Instructions – Report Generator System

This document explains how the Report Generator works, how data flows through the system, how computed fields operate, how report numbers are generated, and how all generated report data is stored in SQLite.

It is intended for:
- Non‑technical users configuring mappings
- Developers maintaining or extending the system

---

# 1. Overview

The Report Generator creates Word `.docx` reports using:
- An Excel `.xlsx` file as input data
- A Word template containing placeholders (`{{placeholder}}`)
- A `mapping.json` file describing how Excel fields map to placeholders
- Optional computed fields (automatic values)
- A SQLite database storing all generated report metadata

The system supports:
- Automatic date fields
- Text transformations
- Custom formatting
- Concatenation
- Automatic report numbering
- Full JSON storage of all used data

---

# 2. Data Flow

```
Excel (.xlsx)
     ↓
Mapping.json (rules)
     ↓
Computed Fields (operations)
     ↓
Report Number Generator
     ↓
Word Template (.docx)
     ↓
Generated Report (.docx)
     ↓
SQLite Database (audit log)
```

---

# 3. Mapping Configuration (`mapping.json`)

This file defines how Excel columns and computed fields map to Word placeholders.

## 3.1 Simple Excel → Placeholder mapping

```json
{
  "date": "{{date}}",
  "species": "{{species}}",
  "address": "{{address}}"
}
```

Each key corresponds to an Excel column name.  
Each value corresponds to a placeholder inside the Word template.

---

# 4. Computed Fields

Computed fields allow generating values that do not exist in Excel.

A computed field is defined as:

```json
"field_name": {
  "operation": "operation_name",
  "...": "parameters"
}
```

The placeholder in the Word template must match the field name:

```
{{field_name}}
```

---

# 5. Supported Operations

## 5.1 `today`
Generates today’s date.

```json
"today_date": {
  "operation": "today",
  "format": "%d/%m/%Y"
}
```

## 5.2 `uppercase`
Converts an Excel field to uppercase.

```json
"species_upper": {
  "operation": "uppercase",
  "input": "species"
}
```

## 5.3 `lowercase`
Converts an Excel field to lowercase.

```json
"species_lower": {
  "operation": "lowercase",
  "input": "species"
}
```

## 5.4 `format`
Formats a string using Excel values.

```json
"full_line": {
  "operation": "format",
  "format": "{species} inspected on {date}"
}
```

## 5.5 `concat`
Concatenates multiple Excel fields.

```json
"combined": {
  "operation": "concat",
  "parts": ["species", "address"]
}
```

## 5.6 `report_number`
Generates a unique report number based on the date and a daily counter.

Format:

```
YYMMDD-XX
```

Example:

```
260611-03
```

Configuration:

```json
"report_number": {
  "operation": "report_number"
}
```

---

# 6. Report Number Generation

The system stores the last generated report number in:

```
/mappings/report_state.json
```

Example:

```json
{
  "last_report_number": "260611-02"
}
```

Logic:
1. Compute today’s prefix: `YYMMDD`
2. If the last number starts with the same prefix:
   - Increment the counter
3. Otherwise:
   - Reset counter to `01`
4. Save the new number back to `report_state.json`

---

# 7. Word Template Requirements

Placeholders must be written as:

```
{{placeholder_name}}
```

Examples:

```
{{date}}
{{species}}
{{today_date}}
{{report_number}}
{{full_line}}
```

---

# 8. SQLite Database Storage

All generated reports are stored in:

```
/mappings/reports.db
```

## 8.1 Table Structure

```
reports (
    id TEXT PRIMARY KEY,          -- report_number
    created_at TEXT,
    excel_path TEXT,
    template_path TEXT,
    row_number INTEGER,
    data_json TEXT                -- full JSON of all data used
)
```

## 8.2 What is stored

- The report number (as primary key)
- The Excel file used
- The Word template used
- The row number used
- A JSON dump of:
  - Excel row data
  - All computed fields
  - All mapping results

This creates a complete audit trail of every generated report.

---

# 9. Developer Architecture

```
app/
  core/
    report_generator.py     → orchestrates everything
    processors.py           → all operations (today, uppercase…)
    state_manager.py        → manages report_state.json
    report_repository.py    → SQLite storage
    mapping_loader.py       → loads mapping.json
    excel_reader.py         → reads Excel
    word_processor.py       → fills Word template
```

---

# 10. Adding New Operations

To add a new computed field operation:

1. Add a function in `processors.py`
2. Register it in `report_generator.py` under `self.operations`
3. Use it in `mapping.json`

Example:

```json
"custom_field": {
  "operation": "my_new_operation",
  "input": "species"
}
```

---

# 11. Summary

- All configuration is done in `mapping.json`
- Computed fields allow dynamic values
- Report numbers are generated automatically and stored
- All report metadata is saved in SQLite
- The system is modular and easy to extend

This document should be used as the reference for configuring and maintaining the Report Generator system.

