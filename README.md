# Rapport-Generator

A desktop application for generating Word `.docx` reports from Excel data, driven by JSON or XLSX mapping configurations and a Tkinter GUI. All report-state and UI state are persisted in a local SQLite database.

---

## Features

- Select a mapping file, Excel data file, and Word template from the GUI
- Supports mapping files in both `.json` and `.xlsx` formats
- `.xlsx` mappings use a `mappings` sheet and may include an optional `config` or `_config` sheet
- Selecting a mapping auto-fills the Excel and template paths from its `config` block
- Manually overriding Excel or template paths writes the change back into the mapping file
- Generate reports with a **Save As** dialog — pre-filled with the suggested folder and filename
- The chosen output folder is persisted to both the mapping file and SQLite for next time
- Support for multiple mapping configurations (one per report type / template)
- Computed fields: today's date, uppercase/lowercase, string formatting, concatenation, auto-incrementing report numbers
- New field transformations: numeric scaling, rounding, date formatting, suffix/prefix, and Excel formula evaluation
- Report counter stored in SQLite and auto-migrated from legacy `report_state.json` if present
- Mapping paths are registered in SQLite so the last used mapping is known and can be reused
- Full audit trail: every generated report is logged to SQLite
- All GUI state (paths, last row) restored automatically on next launch

---

## Requirements

- Python 3.10+

```text
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

## Testing

```powershell
# Activate the virtual environment first
.\env\Scripts\Activate.ps1
python -m pytest
```

If the virtual environment is not yet installed, run:

```powershell
python -m venv env
.\env\Scripts\Activate.ps1
pip install -r requirements.txt
```

---

## Project structure

```text
Rapport-Generator/
├── run.py                              # Convenience entry point
├── requirements.txt
├── app_data.db                         # SQLite database (auto-created on first run)
└── app/
    ├── app.py                          # Application entry point
    ├── core/
    │   ├── excel_reader.py             # Reads rows from .xlsx
    │   ├── mapping_loader.py           # Loads and updates .json/.xlsx mappings
    │   ├── processors.py               # Built-in field operations
    │   ├── report_generator.py         # Orchestrates report creation
    │   ├── state_manager.py            # Report counter migration and generation
    │   └── word_processor.py           # Fills {{placeholders}} in .docx templates
    ├── gui/
    │   ├── config_persistence.py       # GUI state save/restore via SQLite
    │   ├── file_dialogs.py             # File picker and Save As helpers
    │   ├── line_selector.py            # [ − ] [ row ] [ + ] widget
    │   ├── main_window.py              # Main Tkinter window layout and wiring
    │   └── report_actions.py           # Generation logic and output path resolution
    └── repository/
        ├── config_repository.py        # SQLite GUI config + mappings registry
        └── rapport_repository.py       # SQLite report log + state
```

---

## GUI workflow

### On launch
1. Last mapping is restored → its `config` block auto-fills the Excel and template paths
2. Last row number is restored as `last_line_number + 1`
3. Last output folder is remembered for the Save As dialog

### Generating a report
1. Select a **mapping file** → Excel and template paths are filled automatically
2. Adjust the **row number** with `[ − ]` / `[ + ]` or type directly
3. Click **Generate report**
4. A **Save As** dialog opens, pre-filled with the suggested folder and filename
5. Confirm or change the location → report is saved there
6. If the output folder changed, it is written back to both the mapping file and SQLite
7. Row number auto-increments for the next report

---

## Mapping configuration (`mapping.json` / `mapping.xlsx`)

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

### JSON mapping files

Use a JSON file when you want a plain text mapping configuration.

### XLSX mapping files

XLSX mappings must include a sheet named `mappings`.
The first row must define headers such as `Spreadsheet Column`, `Placeholder`, and `Operation`.
Computed rows start with `(computed)` in the `Spreadsheet Column` cell.

A `.xlsx` mapping can also include a `config` or `_config` sheet with key/value pairs.
That sheet may define the same `data_file`, `template_file`, `output_dir`, and additional metadata.

### Simple column mapping

Maps an Excel column header to a `{{placeholder}}` in the Word template:

```json
"species": "{{species}}"
```

The key must match the Excel column header exactly (case-sensitive, trimmed).

### Column mapping with operations

Use transformation rules for fields that need processing before insertion into the document.

```json
"imp": {
  "column": "imp",
  "placeholder": "{{imp}}",
  "operations": [
    { "type": "formula" },
    { "type": "multiply", "value": 100 },
    { "type": "round", "decimals": 0 },
    { "type": "suffix", "value": "%" }
  ]
},
"k value": {
  "column": "k value",
  "placeholder": "{{k_value}}",
  "operations": [
    { "type": "formula" },
    { "type": "multiply", "value": 100 },
    { "type": "round", "decimals": 0 },
    { "type": "suffix", "value": "%" }
  ]
}
```

In this format:
- `column` is the Excel header name
- `placeholder` is the template token written in the Word document
- `operations` is an ordered list of transformations applied to the value

Use this for any future column that needs a calculation, formatting step, or formula-based value.

### Computed fields

```json
"field_name": {
  "operation": "<operation_name>",
  ...options...
}
```

The key becomes the placeholder name (`{{field_name}}`).

---

## Supported computed-field operations

  | Operation         | Description                                                                     | Parameters                                    |
  |-------------------|---------------------------------------------------------------------------------|-----------------------------------------------|
  | `today`           | Today's date                                                                    | `format` — strftime string                    |
  | `uppercase`       | Excel field value in uppercase                                                  | `input` — Excel column or field name          |
  | `lowercase`       | Excel field value in lowercase                                                  | `input` — Excel column or field name          |
  | `format`          | Python-style string template using any available field                         | `format` — e.g. `"{species}_{date}"`         |
  | `concat`          | Concatenate several values from Excel columns or computed fields               | `parts` — list of names                        |
  | `report_number`   | Auto-incrementing report ID for the day (`YYMMDD-XX`)                          | _(none)_                                      |
  | `report_prefix`   | Returns a prefix string from mapping config                                    | _(none)_                                      |
  | `excel_day_counter` | Generates a daily counter based on a date column                             | `date_column`, `date_format`                  |
  | `lookup`          | Simple foreign-key lookup across Excel sheets                                  | `sheet`, `key`, `match`, `value`              |
  | `lookup_join`     | Multi-step relational lookup (join chain)                                      | `steps`, `format`                              |

  ### Supported field-transformation operations

  Use these inside a mapping `operations` list when you need to transform a value from an Excel column before placing it in the document.

  | Type        | Description                                                         | Parameters                              |
  |-------------|---------------------------------------------------------------------|-----------------------------------------|
  | `formula`    | No-op; reads the evaluated Excel value when the workbook is loaded with `data_only=True` | _(none)_                          |
  | `multiply`   | Multiply the value by a number                                      | `value`                                 |
  | `divide`     | Divide the value by a number                                        | `value`                                 |
  | `add`        | Add a number to the value                                           | `value`                                 |
  | `subtract`   | Subtract a number from the value                                    | `value`                                 |
  | `round`      | Round a numeric value                                                | `decimals`                              |
  | `date_format`| Parse and render a date value in a specified format                  | `format`                                |
  | `suffix`     | Append text after the value                                         | `value`                                 |
  | `prefix`     | Prepend text before the value                                       | `value`                                 |
  | `upper`      | Uppercase the text                                                  | _(none)_                                |
  | `lower`      | Lowercase the text                                                  | _(none)_                                |
  | `strip`      | Trim whitespace from the text                                       | _(none)_                                |

  ### `format` field references

  The `format` operation can reference:
  - Any Excel column by its header name: `{species}`, `{date}`
  - Any previously computed field: `{report_number}`, `{today_date}`
  - Fields available from column mappings with `operations` if they are also stored in the mapping result
  - `file_name` is always resolved last so it can use `{report_number}`

  ### Examples

  ```json
  "today_date":     { "operation": "today",            "format": "%d/%m/%Y" },
  "species_upper":  { "operation": "uppercase",        "input": "species" },
  "label":          { "operation": "format",           "format": "{species} – {address}" },
  "ref":            { "operation": "concat",           "parts": ["inspector", "species"] },
  "report_number":  { "operation": "excel_day_counter", "date_column": "date rapport", "date_format": "%d/%m/%Y" },
  "file_name":      { "operation": "format",           "format": "{species}_{date}_{report_number}.docx" }
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

  Note: the app also populates `{{numéro_rapport}}` and `{{sample_number}}` with the generated report number.

  ---

  ## Report numbering

  Format: `YYMMDD-XX` — e.g. `260627-03` is the third report on 27 June 2026.

  The counter resets to `01` each new day. It is stored in the `report_state` SQLite table and updated on every successful generation.

  **Migration:** if a `report_state.json` exists from a previous version, its value is imported into SQLite automatically on first run and the file is deleted.

  ---

  ## Multiple mappings

  Place mapping files anywhere and name them freely:

  ```
  my-mappings/
      animals.json
      inspection.xlsx
      vehicles.json
      custom.xlsx
  ```

  Click **Select mapping file** in the GUI to switch between them. The last selected mapping is remembered across sessions and tracked in SQLite.

  ---

  ## SQLite database (`app_data.db`)

  Created automatically in the project root on first run. Contains four tables:

  ### `config` — GUI persistent state

  | Key                | Description                                      |
  |--------------------|--------------------------------------------------|
  | `last_excel_path`  | Last selected `.xlsx` path                       |
  | `last_docx_path`   | Last selected `.docx` template path              |
  | `mapping_path`     | Last selected mapping file path                  |
  | `last_line_number` | Last used row number                             |
  | `last_output_dir`  | Last output folder chosen in Save As             |

  > Legacy compatibility: `last_mapping_path` is still read when present, but `mapping_path` is the preferred key.

  ### `mappings` — mapping file registry

  | Column         | Description                |
  |----------------|----------------------------|
  | `id`           | Auto-increment identifier  |
  | `mapping_path` | Stored mapping file path   |

  This table tracks all mappings selected through the GUI.

  ### `reports` — audit log

  | Column          | Type    | Description                           |
  |-----------------|---------|---------------------------------------|
  | `id`            | TEXT    | Unique ID (`row_N_YYYYMMDDHHMMSS`)    |
  | `created_at`    | TEXT    | ISO 8601 timestamp                    |
  | `excel_path`    | TEXT    | Path to the Excel source              |
  | `template_path` | TEXT    | Path to the Word template             |
  | `mapping_path`  | TEXT    | Path to the mapping file              |
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
  report_state.json
  .DS_Store
  Thumbs.db
  .vscode/
  .idea/
  .pytest_cache/
  ```