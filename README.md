# Rapport-Generator

A desktop application for generating Word `.docx` reports from Excel data, driven by JSON mapping configurations and a Tkinter GUI. All state is persisted in a local SQLite database — no loose JSON files.

---

## Features

- Select a mapping file, Excel data file, and Word template from the GUI
- Selecting a mapping auto-fills the Excel and template paths from its `config` block
- Manually overriding Excel or template paths writes the change back into the mapping JSON
- Generate reports with a **Save As** dialog — pre-filled with the suggested folder and filename
- The chosen output folder is persisted to both the mapping JSON and SQLite for next time
- Support for multiple mapping configurations (one per report type / template)
- Computed fields: today's date, uppercase/lowercase, string formatting, concatenation, auto-incrementing report numbers
- Report counter stored in SQLite — no `report_state.json` file; auto-migrated from JSON on first run if one exists
- Full audit trail: every generated report is logged to SQLite
- All GUI state (paths, last row) restored automatically on next launch

---

## Requirements

- Python 3.10+

```
et_xmlfile==2.0.0
lxml==6.1.1
openpyxl==3.1.5
python-docx==1.2.0
typing_extensions==4.15.0
```

---

## Installation

```powershell
python -m venv env
.\env\Scripts\Activate.ps1
pip install -r requirements.txt
```

---

## Running

```powershell
python run.py
```

Or directly:

```powershell
python app/app.py
```

---

## Project structure

```
Rapport-Generator/
├── run.py                              # Convenience entry point
├── requirements.txt
├── app_data.db                         # SQLite database (auto-created on first run)
└── app/
    ├── app.py                          # Application entry point
    ├── core/
    │   ├── excel_reader.py             # Reads rows from .xlsx
    │   ├── mapping_loader.py           # Loads, queries, and updates mapping.json
    │   ├── processors.py               # Built-in field operations (today, format, …)
    │   ├── report_generator.py         # Orchestrates report creation
    │   ├── state_manager.py            # Report number counter (SQLite-backed)
    │   └── word_processor.py           # Fills {{placeholders}} in .docx templates
    ├── gui/
    │   ├── config_persistence.py       # GUI state save/restore via SQLite
    │   ├── file_dialogs.py             # File picker and Save As helpers
    │   ├── line_selector.py            # [ − ] [ row ] [ + ] widget
    │   ├── main_window.py              # Main Tkinter window — layout and wiring only
    │   └── report_actions.py           # Generation logic and output path resolution
    ├── mappings/
    │   └── data.json                   # Default mapping configuration
    └── repository/
        ├── config_repository.py        # SQLite key/value store for GUI config
        └── rapport_repository.py       # SQLite report log + report state
```

---

## GUI workflow

### On launch
1. Last mapping is restored → its `config` block auto-fills the Excel and template paths
2. Last row number is restored as `last_row + 1`
3. Last output folder is remembered for the Save As dialog

### Generating a report
1. Select a **mapping file** → Excel and template paths are filled automatically
2. Adjust the **row number** with `[ − ]` / `[ + ]` or type directly
3. Click **Generate report**
4. A **Save As** dialog opens, pre-filled with the suggested folder and filename
5. Confirm or change the location → report is saved there
6. If the output folder changed, it is written back to both the mapping JSON and SQLite
7. Row number auto-increments for the next report

---

## Mapping configuration (`mapping.json`)

Each mapping file ties together one Excel table, one Word template, one output folder, and a set of field rules. Selecting a mapping in the GUI is the only step needed to switch report types.

### Full structure

```json
{
  "config": {
    "data_file":     "C:/Reports/table.xlsx",
    "template_file": "C:/Reports/template.docx",
    "output_dir":    "C:/Reports/Output"
  },

  "date":      "{{date}}",
  "species":   "{{species}}",
  "address":   "{{address}}",
  "inspector": "{{inspector}}",
  "notes":     "{{notes}}",
  "product":   "{{product}}",

  "today_date": {
    "operation": "today",
    "format": "%d/%m/%Y"
  },
  "species_upper": {
    "operation": "uppercase",
    "input": "species"
  },
  "full_line": {
    "operation": "format",
    "format": "{species} inspected on {date}"
  },
  "report_number": {
    "operation": "report_number"
  },
  "file_name": {
    "operation": "format",
    "format": "{species}_{date}_{report_number}.docx"
  }
}
```

### `config` block

| Key             | Description                                          |
|-----------------|------------------------------------------------------|
| `data_file`     | Path to the Excel `.xlsx` data file                  |
| `template_file` | Path to the Word `.docx` template                    |
| `output_dir`    | Default output folder for the Save As dialog         |

All three are updated automatically when the user selects different files or a different output folder in the GUI.

### Simple column mapping

Maps an Excel column header to a `{{placeholder}}` in the Word template:

```json
"species": "{{species}}"
```

The key must match the Excel column header exactly (case-sensitive, trimmed).

### Computed fields

```json
"field_name": {
  "operation": "<operation_name>",
  ...options...
}
```

The key becomes the placeholder name (`{{field_name}}`).

---

## Supported operations

| Operation       | Description                                                   | Parameters                                    |
|-----------------|---------------------------------------------------------------|-----------------------------------------------|
| `today`         | Today's date                                                  | `format` — strftime string                    |
| `uppercase`     | Excel field value in uppercase                                | `input` — Excel column name                   |
| `lowercase`     | Excel field value in lowercase                                | `input` — Excel column name                   |
| `format`        | Python-style format string using any column or computed field | `format` — e.g. `"{species}_{date}"`          |
| `concat`        | Concatenate multiple Excel columns                            | `parts` — list of column names                |
| `report_number` | Auto-incrementing daily counter (`YYMMDD-XX`)                 | _(none)_                                      |

### `format` field references

The `format` operation can reference:
- Any Excel column by its header name: `{species}`, `{date}`
- Any previously computed field: `{report_number}`, `{today_date}`
- `file_name` is always resolved last so it can use `{report_number}`

### Examples

```json
"today_date":     { "operation": "today",     "format": "%d/%m/%Y" },
"species_upper":  { "operation": "uppercase",  "input": "species" },
"label":          { "operation": "format",     "format": "{species} – {address}" },
"ref":            { "operation": "concat",     "parts": ["inspector", "species"] },
"report_number":  { "operation": "report_number" },
"file_name":      { "operation": "format",     "format": "{species}_{date}_{report_number}.docx" }
```

---

## Word template placeholders

Use double-brace syntax anywhere in the `.docx` template:

```
{{date}}          {{species}}        {{address}}
{{inspector}}     {{notes}}          {{product}}
{{today_date}}    {{species_upper}}  {{full_line}}
{{report_number}}
```

Any placeholder not matched by the mapping is left unchanged.

---

## Report numbering

Format: `YYMMDD-XX` — e.g. `260627-03` is the third report on 27 June 2026.

The counter resets to `01` each new day. It is stored in the `report_state` SQLite table and updated on every successful generation.

**Migration:** if a `report_state.json` exists from a previous version, its value is imported into SQLite automatically on first run and the file is deleted.

---

## Multiple mappings

Place mapping files anywhere and name them freely:

```
mappings/
    animals.json
    inspection.json
    vehicles.json
    custom.json
```

Click **Select mapping.json** in the GUI to switch between them. The last selected mapping is remembered across sessions.

---

## SQLite database (`app_data.db`)

Created automatically in the project root on first run. Contains three tables:

### `config` — GUI persistent state

| Key                 | Description                                  |
|---------------------|----------------------------------------------|
| `last_excel_path`   | Last selected `.xlsx` path                   |
| `last_docx_path`    | Last selected `.docx` template path          |
| `last_mapping_path` | Last selected mapping `.json` path           |
| `last_line_number`  | Last used row number                         |
| `last_output_dir`   | Last output folder chosen in Save As         |

### `reports` — audit log

| Column          | Type    | Description                           |
|-----------------|---------|---------------------------------------|
| `id`            | TEXT    | Unique ID (`row_N_YYYYMMDDHHMMSS`)    |
| `created_at`    | TEXT    | ISO 8601 timestamp                    |
| `excel_path`    | TEXT    |                                       |
| `template_path` | TEXT    |                                       |
| `mapping_path`  | TEXT    |                                       |
| `row_number`    | INTEGER | Excel row used                        |
| `data_json`     | TEXT    | Full data snapshot (JSON)             |

### `report_state` — report counter

| Key                  | Value             |
|----------------------|-------------------|
| `last_report_number` | e.g. `260627-03`  |

---

## Adding a new operation

1. Add a function to `app/core/processors.py`:

```python
def op_titlecase(rule, row_data):
    field = rule.get("input", "")
    return str(row_data.get(field, "")).title()
```

2. Register it in `app/core/report_generator.py`:

```python
self.operations = {
    ...
    "titlecase": processors.op_titlecase,
}
```

3. Use it in a mapping file:

```json
"species_title": {
  "operation": "titlecase",
  "input": "species"
}
```

4. Add `{{species_title}}` to the Word template.

---

## .gitignore

```gitignore
env/
.venv/
__pycache__/
*.py[cod]
*.db
*.sqlite
report_row_*.docx
app/mappings/report_state.json
.DS_Store
Thumbs.db
.vscode/
.idea/
.pytest_cache/
```