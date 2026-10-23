# Report Generator — User & Technical Guide

<!-- This file is shown in the app (📖 Guide) and copied to the clipboard for AI assistants. -->
<!-- Lines like this one are removed before display/copy. -->

**Version of this guide:** matches Report Generator 1.x (profiles, mapping JSON, Word template, Excel data).

## 0. If you are an AI assistant reading this

The user pasted this guide so you can help them fix or write one of the three files the app needs:
the **mapping** (`mapping.json`), the **Word template** (`template.docx`) or the **Excel data** (`data.xlsx`).

How to help well:

1. **Never invent names.** Column names, sheet names and placeholders must be copied exactly from
   what the user gives you (the Excel headers, the placeholders found in the template, the warnings
   shown by the app). Names are case-sensitive and must match character for character.
2. **Return the complete corrected JSON**, not only a fragment, unless the user asks otherwise.
   Keep every rule you were not asked to change. The JSON must be valid (double quotes, commas,
   no comments, no trailing comma).
3. **Explain each change in one line** (what was wrong → what you changed → why).
4. If something is missing (for example you do not know the Excel headers), **ask for it** instead
   of guessing. The user can paste the "Excel headers" and "Template placeholders" shown in the
   profile context.
5. Word and Excel cannot be edited by you directly: tell the user precisely what to change
   (which placeholder to retype, which header to rename, which row to add to which sheet).
6. Section 9 explains every warning the app can show, and section 10 lists typical symptoms.
   Start there when the user pastes a warning or a symptom.

## 1. What the app does

Report Generator fills a **Word template** with the values of **one row of an Excel file**, following
the rules of a **mapping file**. The result is a new `.docx` report.

```text
   Excel data file ──(row N)──┐
                              ├──►  Report Generator  ──►  report.docx
   Mapping file (JSON) ───────┤        (applies rules)
   Word template ─────────────┘
```

A **profile** bundles the three files of one kind of report (for example "Certificate of Analysis"
or "Fish quality") so you can switch between report types with one click.

## 2. Quick start

1. Pick a profile in the **Profile** list (or create one with **New…**).
2. Check the status line under the profile: **✔ Ready** means nothing was found. **⚠** means
   warnings: click **Details…** to read them. **✖** means a blocking problem (a file is missing…):
   fix it first.
3. Choose the **row** to turn into a report in **Line (row ≥ 2)**. Row 1 is the header of the
   Excel sheet, so the first data row is row 2.
4. Check **Output Location** and **Name of the report**.
5. Click **Generate report**, choose where to save the file, done. The row number then moves to the
   next row automatically.

## 3. The main window

| Control | What it does |
|---|---|
| **Profile** list | Switches the whole window to another profile: its mapping, template, data file, output folder and next row. **Ctrl+1 … Ctrl+9** select the first nine profiles. |
| **New…** | Creates a profile: a name and (optionally) the three files. It can start from a copy of an existing profile. |
| **Update…** | Same window for the current profile: rename it, or replace its mapping / template / data file. Leave a field empty to keep the current file. The replaced file is kept as `.bak`. |
| **⋯** menu | Update, Refresh, Rename, Duplicate, Replace data / template / mapping, Open profile folder, Export / Import a profile as `.zip`, Delete. |
| Status line, **Details…** | Result of the automatic checks of the profile (see section 9). |
| **⟳ Refresh** (F5) | Re-reads the three files from disk and re-runs the checks. Use it after editing a file outside the app. |
| **📂 Mapping / Template / Data** | Opens that file of the profile with its default program (Notepad, Word, Excel…). |
| **Output Location** | Folder proposed when saving the report. Remembered per profile. |
| **Name of the report** | The fixed part of the file name (the `name` of the `file_name` rule, see 7.6). |
| **Line (row ≥ 2)** | Excel row used for the report. **−** / **+** change it. Remembered per profile. |
| **Report number: …** | Preview of the report number of the selected row (only when the mapping computes one). |
| **Generate report** | Builds the report. Disabled while the profile has a blocking problem (✖). |
| **📖 Guide** | Opens this guide. You can copy it, or copy it together with your profile to ask an AI for help. |

## 4. Profiles

A profile is a **folder with three files with fixed names**:

```text
<profiles folder>\
    my-profile\
        mapping.json     the rules (section 7)
        template.docx    the Word template (section 6)
        data.xlsx        the Excel data (section 5)
```

- Installed app: `%APPDATA%\ReportGenerator\profiles\`. Running from the source code: the `profiles`
  folder next to `run.py`. **Open profile folder** in the **⋯** menu opens it.
- The files are **copies**: editing the original file you selected at creation has no effect on
  the profile. Edit the profile's own files (the **📂** buttons) or use **Update…** to replace them.
- Each profile remembers its own next row, its own output folder and its own daily report counter.
- Deleting a profile moves its folder to `_trash` inside the profiles folder, it is not destroyed.
- A folder with a `mapping.json` that you copy by hand into the profiles folder appears
  automatically (press **⟳ Refresh**). The display name is read from `config.profile_name`.
- **Export / Import** a profile as a `.zip` to share it with a colleague.

## 5. The Excel data file (`data.xlsx`)

### 5.1 Structure

- The workbook has one **data sheet** plus, optionally, **lookup sheets** (small reference tables).
- The data sheet is `config.data_sheet` in the mapping. If it is not set, the **first sheet** is used.
- **Row 1 = headers** (column names). **Row 2 and below = one report per row.**
- Headers are matched **exactly**: capitals, accents, spaces and hyphens all count.
  `K-value` is not `k value`. Leading and trailing spaces in a header are ignored.
- Empty cells give empty text. Avoid completely empty rows between data rows.
- Write the header once and never rename it without updating the mapping.

### 5.2 Dates and numbers

- Real Excel dates are recommended (cell formatted as a date). Without operations, a date cell is
  written with `config.date_format` (default `%d/%m/%Y`, e.g. `04/03/2026`).
- Numbers are written as plain text (`12.3456`). To control decimals, use the `number_format`
  operation (section 7.4).
- A cell with text such as `12,5` is understood as a number by the numeric operations.

### 5.3 Formulas

- If a cell is a formula and Excel has saved its result, that saved result is used.
- If no saved result exists, the app can compute **simple arithmetic only** (`+ - * / ^ %` and
  references to other cells of the same sheet, e.g. `=A2*B2`). Anything else (`IF`, `SUM`, `VLOOKUP`…)
  gives an **empty value**.
- To be safe: **open the file in Excel and save it** before generating reports, and keep
  complex formulas calculated by Excel.
- After each report the app writes back into the file (see 8.3) with a library that does not keep the
  saved results of formulas. If a formula column suddenly becomes empty after generating, open the
  file in Excel and save it again.

### 5.4 Lookup sheets

A lookup sheet is a table with a header row, used to find a value from another value. Example,
sheet `suppliers`:

| Supplier | Approval number |
|---|---|
| ACME Seafood | FR 62.001.001 CE |
| Boulogne Marée | FR 62.160.025 CE |

The mapping can then find the approval number of the supplier written in the data sheet
(`lookup`, section 7.5). The comparison ignores capitals and surrounding spaces, but not spelling:
`Acme Seafod` will **not** match `ACME Seafood`.

### 5.5 What the app writes into the Excel file

After a report is generated, if the mapping computes a report number, the app writes into the
data sheet, on that row: `numero rapport` (e.g. `260621-3`) and `nom rapport generé` (full report
name). The columns are created if they do not exist.
**Close the Excel file before generating**, otherwise the write-back fails (silently) because Excel
locks the file. Keep a backup of important data files.

## 6. The Word template (`template.docx`)

### 6.1 Placeholders

A placeholder is a name between double braces: `{{sample_name}}`. The app replaces it with the value
of the field of the same name.

Rules (all of them matter):

- Exactly **two opening braces and two closing braces**: `{{name}}`. `{name}}` or `{{name}` do not work.
- **No spaces** inside the braces: `{{ name }}` is **not** replaced.
- **Case-sensitive**: `{{Name}}` and `{{name}}` are different. Accents count (`{{numéro_rapport}}`).
- Use letters, digits and underscores. Avoid spaces and other special characters.
- The same placeholder may appear several times; every occurrence is replaced.
- A placeholder that no rule fills is **left as it is** in the report (you will see `{{name}}` in
  the Word file). That is the visible sign of a mapping or template mistake.

### 6.2 Where placeholders work

| Place | Works? |
|---|---|
| Normal paragraphs | yes |
| Tables, including nested tables | yes |
| **Headers and footers** | **no** — placeholders there are not replaced |
| **Text boxes** and shapes | **no** |
| Content controls, fields, footnotes, hyperlinks | not supported |

If you need a value in a header, repeat the information in the body or in a table.

### 6.3 Formatting

- The formatting (font, size, bold, colour) of the first character of the placeholder is kept for
  the inserted value. Format the placeholder in Word as you want the value to look.
- Word sometimes splits a placeholder into several pieces internally (spell-check, edits). The app
  repairs that automatically, the final formatting is that of the first piece.
- A value that contains a line break (`\n` produced by the `replace` operation) becomes a real line
  break in Word.
- Tip: type the placeholder in one go, without correcting it afterwards. If a placeholder refuses
  to be replaced, delete it and type it again, or paste it as plain text.

### 6.4 Special placeholders

If the mapping defines a field named `sample_number` or `numero_rapport`, its value is also
written to `{{sample_number}}` and `{{numéro_rapport}}` automatically.

## 7. The mapping file (`mapping.json`)

The mapping is a JSON file: a list of **fields**. Each field has a name (the **key**) and a rule.
There is also one special block named `config` and one named `file_name`.

### 7.1 General shape

```json
{
  "config": { "date_format": "%d/%m/%Y" },
  "field_name": { "column": "Excel header", "placeholder": "{{field_name}}" },
  "other_field": { "operation": "today", "format": "%d/%m/%Y" },
  "file_name": { "name": "Report", "operation": "format", "format": "{name} {field_name}.docx" }
}
```

JSON reminders: keys and text use **double quotes**; fields are separated by commas, with **no comma
after the last one**; backslashes must be doubled (`"\\n"`); comments are not allowed.

There are two kinds of fields:

- **Column fields** (7.3): take a value from an Excel column, optionally transform it
  (`column`, `placeholder`, `operations`).
- **Calculated fields** (7.5): compute a value (`operation`, plus parameters).

### 7.2 The `config` block

| Key | Meaning |
|---|---|
| `data_sheet` | Name of the data sheet. Default: the first sheet. |
| `date_format` | How dates are written when a date cell has no `date_format` operation. Default `%d/%m/%Y`. |
| `report_prefix` | Free text available as `{report_prefix}` in formats, and returned by the `report_prefix` operation. |
| `output_dir` | Folder proposed when saving. Updated by the app when you choose another folder. |
| `name` | If present, replaces the `name` of the `file_name` rule. |
| `profile_name` | Display name of the profile (written by the app). |
| `data_file`, `template_file` | Old settings (before profiles). Ignored: a profile always uses its own files. |

Date format codes (Python `strftime`): `%d` day 01-31 · `%m` month 01-12 · `%Y` year 2026 ·
`%y` year 26 · `%B` month name (English) · `%b` short month · `%A` weekday · `%H` hour 00-23 ·
`%M` minutes · `%S` seconds. Example: `%d/%m/%Y` → `04/03/2026`, `%Y-%m-%d` → `2026-03-04`.

### 7.3 Column fields

```json
"end_date": {
  "column": "end date",
  "placeholder": "{{end_date}}",
  "operations": [ { "type": "date_format", "format": "%d/%m/%Y" } ]
}
```

| Property | Required | Meaning |
|---|---|---|
| key (`"end_date"`) | yes | Name of the field. Free choice; used in formats (`{end_date}`). |
| `column` | yes | Excel header of the data sheet, **exact** spelling. |
| `placeholder` | yes | The placeholder in the Word template, with braces: `"{{end_date}}"`. |
| `operations` | no | List of transformations applied in order (7.4). |

Both `column` and `placeholder` are needed, otherwise the field is ignored.
If the column is not found in the Excel file, or the placeholder is not in the template, the field is
silently skipped (the app warns you in the status line).

Without `operations`, the cell value is written as text. With `operations`, the **raw cell value**
(real date, real number) is given to the first operation.

### 7.4 Operations on column fields

Operations run **in order**, each one receiving the result of the previous one.

| `type` | Parameters | Effect | Example → result |
|---|---|---|---|
| `date_format` | `format` | Writes a date with the given format. Accepts a date cell or text like `2026-03-04`, `04/03/2026`. | `"%d %B %Y"` → `04 March 2026` |
| `number_format` | `decimals` (default 2), `decimal_separator` (`.` or `,`), `thousands` (e.g. `" "`) | Fixed number of decimals, half-up rounding (1.005 → 1.01). Empty or non-numeric values are left unchanged. | `12.3456`, decimals 2, separator `,` → `12,35` |
| `round` | `decimals` | Rounds to a number (no trailing zeros; 0 decimals gives a whole number). | `0.4567`, decimals 0 → `0` |
| `multiply` `divide` `add` `subtract` | `value` | Arithmetic with a number. `12,5` (comma) is understood. Non-numeric values are left unchanged. | `0.92` multiply 100 → `92.0` |
| `prefix` / `suffix` | `value` | Adds text before / after. | suffix `"%"` → `92%` |
| `upper` / `lower` | — | Capitals / small letters. | |
| `strip` | — | Removes spaces at both ends. | |
| `replace` | `find`, `replace` | Replaces text. Use `"\\n"` (in JSON) as replacement to insert a line break. | find `", "` replace `"\\n"` |
| `formula` | — | Does nothing (marker: the cell is a formula whose result is read). | |

Examples:

```json
"operations": [
  { "type": "multiply", "value": 100 },
  { "type": "round", "decimals": 0 },
  { "type": "suffix", "value": "%" }
]
```
gives `0.9234` → `92%`.

```json
"operations": [
  { "type": "replace", "find": ", ", "replace": "\\n" }
]
```
gives `12 Quai Gambetta, 62200 Boulogne` on two lines in Word.

Important:

- A **misspelled `type` is silently ignored** (`number_fromat` does nothing). The app warns about it.
- A text result after `suffix`/`prefix`/`number_format` cannot be used in later arithmetic. Put
  arithmetic first, formatting last.
- A date cell with operations other than `date_format` first would be written as a raw date-time.
  Put `date_format` first for dates.
- `decimals` must be a whole number, `value` of arithmetic operations a number (no quotes needed;
  `"value": 100`).

### 7.5 Calculated fields

A calculated field has an `operation` instead of a `column`.

```json
"report_date": { "operation": "today", "format": "%d/%m/%Y" }
```

> **Very important:** for a calculated field the placeholder in the Word template is always the
> **key** of the field: `report_date` → `{{report_date}}`. A `placeholder` property, if present, is
> ignored. So the key must be exactly the name you typed in the template.

| `operation` | Parameters | Result |
|---|---|---|
| `today` | `format` (date format, default `%Y-%m-%d`) | Today's date. |
| `uppercase` / `lowercase` | `input` (a column name or a field name) | That value in capitals / small letters. |
| `format` | `format` (text with `{names}`) | A text built from other values, see below. |
| `concat` | `parts` (list of names) | The values stuck together without separator. |
| `report_prefix` | — | `config.report_prefix`. |
| `report_number` | — | Daily counter `yymmdd-NN` (e.g. `260621-03`), one counter per profile. A new number is consumed at each generated report. |
| `excel_day_counter` | `date_column`, `sample_column` (optional), `date_format` | A report number built from the date of the row (see below). |
| `lookup` | `sheet`, `key`, `match`, `value` | Finds a value in a lookup sheet (see below). |
| `lookup_join` | `steps` (list), `format` (optional) | Chains several lookups (see below). |

**Using names inside `format` and `concat`**: write `{name}` where `name` is a field key from the mapping,
an Excel column name (`{sample name}` works, spaces allowed), `{report_prefix}`, or a property of the
rule itself (e.g. `{name}` in `file_name`). Calculated fields can use other calculated fields: the order in the
file does not matter.

> Note: a **column field** can be used as `{key}` only if its placeholder exists in the template.
> Otherwise write the Excel column name instead (`{sample name}`).

```json
"title": { "operation": "format", "format": "{report_prefix} - {sample_name}" }
```

**`excel_day_counter`** builds `yymmdd-X` from the date of the row. With `sample_column`, X is the value of
that column (`260621-A12`). Without it, X is how many rows up to the current one have the same date
(`260621-3`). `date_format` must be the same as `config.date_format`.

```json
"numero_rapport": {
  "operation": "excel_day_counter",
  "date_column": "date rapport",
  "sample_column": "numero echantillon",
  "date_format": "%d/%m/%Y"
}
```

**`lookup`**: for the row, take the value of the data-sheet column `match`, find the row of the sheet
`sheet` whose column `key` has the same value (capitals and surrounding spaces ignored), and return the
column `value` of that row.

```json
"approval": {
  "operation": "lookup",
  "sheet": "suppliers",
  "key": "Supplier",
  "match": "supplier",
  "value": "Approval number"
}
```

If nothing is found, **the placeholder is not replaced** and stays visible as `{{approval}}` in the report.
If the row is found but the looked-up cell is empty, the value is empty.

**`lookup_join`**: a chain. Each step is a lookup; `"match": "<previous>"` means "use the result of the
previous step". `value` can be a list of columns (then use `format` to assemble them).

```json
"quality_criteria": {
  "operation": "lookup_join",
  "steps": [
    { "sheet": "species", "key": "Latin name", "match": "species", "value": "French name" },
    { "sheet": "species-criteria", "key": "Species", "match": "<previous>", "value": "Criteria" }
  ]
}
```

### 7.6 The `file_name` rule

```json
"file_name": {
  "name": "Certificate",
  "operation": "format",
  "format": "{name} {sample_name}.docx"
}
```

- `format` builds the **file name** proposed when saving. It may use `{name}` (the text above),
  field keys, Excel column names and `{report_prefix}`. End it with `.docx`.
- `name` is the editable **Name of the report** in the main window.
- Characters not allowed in Windows file names (`< > : " / \ | ? *`) are replaced by `_`.
- If the format uses a name that does not exist, the file is called `report_row_N.docx`.

### 7.7 Older shorthand

An older form is still read: `"Excel header": "{{placeholder}}"` (the key is the Excel header and the value
is the placeholder). Prefer the full form of 7.3 in new mappings.

## 8. How a report is generated

### 8.1 Step by step

1. The app reads the selected **row** of the data sheet and turns every cell into text (dates with `config.date_format`).
2. It lists the **placeholders** found in the template.
3. **Column fields**: for each one, read the column, apply the operations, and if the placeholder is in the
   template, remember the value.
4. **Calculated fields**: computed in as many passes as needed (a field can use another one).
5. The **file name** is built from the `file_name` rule.
6. The template is copied and every placeholder that has a value is replaced. Other placeholders remain.
7. The report is saved where you chose. The data file is updated (8.3), and the next row is selected.

### 8.2 What is left unchanged on purpose

A placeholder stays as `{{name}}` in the report when: no field produces it, its column does not exist,
its lookup finds nothing, or the field's operations could not be computed. Search the finished Word file
for `{{` to find them quickly.

### 8.3 What is remembered

- **Per profile**: next row, output folder, daily counter (`report_number`).
- **In the data file**: if a report number is computed, the columns `numero rapport` and
  `nom rapport generé` are filled for that row.

## 9. Messages of the automatic check (status line → Details…)

The check runs when you select a profile, after **Update…**, and with **⟳ Refresh**.
**✖ errors** block *Generate report*. **⚠ warnings** do not, but the report will probably be wrong.

| Message | Meaning | How to fix |
|---|---|---|
| `mapping.json / template.docx / data.xlsx is missing` | A file of the profile is absent. | **Update…** and choose the file. |
| `mapping.json is not valid: …` | The JSON has a syntax error (comma, quote, bracket). | Fix at the line given; paste it into an AI or a JSON validator. |
| `Data sheet 'X' not found` | `config.data_sheet` names a sheet that does not exist. | Correct the name (see the sheet tabs in Excel). |
| `Field 'f': column 'c' not found in the data sheet. Did you mean 'C'?` | The `column` is not an exact Excel header. | Copy the header exactly (capitals, hyphens, spaces), or rename the header in Excel. |
| `Field 'f': placeholder {{p}} not found in the template.` | The field writes to a placeholder the template does not contain. | Add `{{p}}` to the template, or change the placeholder in the mapping. |
| `Template placeholder {{p}} is not filled by any field in the mapping` | The template has `{{p}}` but no field produces it; it will stay as `{{p}}` in the report. | Add a field for it, or remove it from the template. |
| `Field 'f': unknown type 'x' — it is ignored` | Misspelled operation in `operations`. | Use one of the types of 7.4. |
| `Field 'f': unknown operation 'x'` | Misspelled calculated operation. | Use one of the operations of 7.5. |
| `… for calculated fields the placeholder in the template is always {{f}}` | A calculated field's key differs from the placeholder you wrote. | Rename the field's key to the placeholder name (or change the template). |
| `Field 'f' has neither a 'column' nor an 'operation'` | The rule does nothing. | Add a `column` or an `operation`. |
| `'decimals' must be a whole number` / `'value' must be a number` | Wrong parameter type. | `"decimals": 2` (no quotes). |
| `file_name: '{x}' is not a field name or data column` | The file name format uses an unknown name. | Use a field key or an Excel column name. |
| `lookup sheet 'S' not found in the data file` / `column 'C' not found in sheet 'S'` | The lookup points to a sheet or column that does not exist. | Fix the name; the check suggests the closest one. |
| `N value(s) of 'col' have no match in sheet 'S' … e.g. 'abc' (row 4)` | Those cells have no equivalent in the lookup sheet: the placeholder stays as `{{field}}` for those rows. | Add the missing entries to the lookup sheet, or correct the spelling in the data (the check suggests close names). |
| `… are found but the looked-up cell is empty` | The lookup sheet has the entry but its value cell is empty. | Fill the cell in the lookup sheet. |

## 10. Troubleshooting by symptom

| Symptom | Likely cause | What to do |
|---|---|---|
| `{{something}}` still visible in the report | No field fills it / column not found / lookup found nothing (see 8.2) | Read **Details…**, fix the matching item of section 9. |
| A value is missing but the placeholder disappeared | The Excel cell is empty, or a formula had no saved result | Check the cell; open and save the file in Excel (5.3). |
| Placeholder in the template is never replaced although the mapping looks right | Spaces inside the braces, wrong capitals, one brace missing, or it is in a header/footer/text box | Retype it as `{{name}}` in the body or a table (section 6). |
| File is called `report_row_N.docx` | The `file_name` format uses an unknown name | Fix `file_name.format` (7.6). |
| Date is shown as `2026-03-04 00:00:00` or in the wrong format | An operation other than `date_format` came first | Put `{ "type": "date_format", … }` first, or set `config.date_format`. |
| Number shows too many decimals | No `number_format` | Add `{ "type": "number_format", "decimals": 2 }`. |
| Month names are in English | `%B` uses English | Use a lookup sheet for French names, or a `replace` operation. |
| `Row N out of range` | The row number is below 2 or beyond the last data row | Choose an existing row. |
| Report number differs from what you expect | `report_number` is a counter per day and per profile; `excel_day_counter` depends on `date_column`, `sample_column` and the date format | Check those parameters (7.5). |
| Nothing was written back into the Excel file | The file was open in Excel (locked), or the mapping has no report number | Close Excel; the write-back needs `sample_number` or `numero_rapport`. |
| After a report, a formula column became empty | Saving by the app drops the saved results of formulas | Open the file in Excel and save it (5.3). |
| The mapping edited outside the app is not taken into account | The window shows the previous state | Click **⟳ Refresh** (F5). |
| **Update…** says `… is open in another program (Word or Excel)` | The profile's file is open (for example after clicking **📂 Data**), and Windows does not allow replacing an open file | Close Word/Excel, then Save again. Nothing was changed. If you already edited the profile's own file (opened with **📂**), you do not need **Update…**: save it, close it and press **⟳ Refresh**. |

## 11. A complete worked example

Data file, sheet `Data`:

| sample name | reception date | weight | supplier | address |
|---|---|---|---|---|
| Sample A1 | 2026-03-04 (a date cell) | 12.3456 | acme seafood | 12 Quai Gambetta, 62200 Boulogne-sur-Mer |

Sheet `suppliers`:

| Supplier | Approval number |
|---|---|
| ACME Seafood | FR 62.001.001 CE |

Word template (placeholders typed as plain text):

```text
CERTIFICATE {{title}}
Sample: {{sample_name}}
Received on: {{reception_date}}
Weight: {{weight}} g
Supplier approval: {{approval}}
Address:
{{address}}
Issued on {{report_date}}
```

Mapping:

<!-- test:example-mapping -->
```json
{
  "config": {
    "data_sheet": "Data",
    "date_format": "%d/%m/%Y",
    "report_prefix": "COA"
  },
  "sample_name": {
    "column": "sample name",
    "placeholder": "{{sample_name}}"
  },
  "reception_date": {
    "column": "reception date",
    "placeholder": "{{reception_date}}",
    "operations": [
      { "type": "date_format", "format": "%d/%m/%Y" }
    ]
  },
  "weight": {
    "column": "weight",
    "placeholder": "{{weight}}",
    "operations": [
      { "type": "number_format", "decimals": 2, "decimal_separator": "," }
    ]
  },
  "address": {
    "column": "address",
    "placeholder": "{{address}}",
    "operations": [
      { "type": "replace", "find": ", ", "replace": "\\n" }
    ]
  },
  "approval": {
    "operation": "lookup",
    "sheet": "suppliers",
    "key": "Supplier",
    "match": "supplier",
    "value": "Approval number"
  },
  "report_date": {
    "operation": "today",
    "format": "%d/%m/%Y"
  },
  "title": {
    "operation": "format",
    "format": "{report_prefix} - {sample_name}"
  },
  "file_name": {
    "name": "Certificate",
    "operation": "format",
    "format": "{name} {sample_name}.docx"
  }
}
```

Result of the report for row 2:

```text
CERTIFICATE COA - Sample A1
Sample: Sample A1
Received on: 04/03/2026
Weight: 12,35 g
Supplier approval: FR 62.001.001 CE
Address:
12 Quai Gambetta
62200 Boulogne-sur-Mer
Issued on <today's date>
```

File name proposed: `Certificate Sample A1.docx`.

## 12. Cheat sheet

- Excel: headers in row 1, data from row 2, exact spelling, save in Excel before generating, close Excel while generating.
- Word: `{{name}}`, no spaces, same capitals, in the body or tables (not headers/footers/text boxes).
- Column field: `"key": { "column": "Excel header", "placeholder": "{{name}}", "operations": [ … ] }`
- Calculated field: `"name": { "operation": "…", … }` — the key **is** the placeholder name.
- Operations: `date_format` `number_format` `round` `multiply` `divide` `add` `subtract` `prefix` `suffix` `upper` `lower` `strip` `replace` `formula`.
- Calculated operations: `today` `uppercase` `lowercase` `format` `concat` `report_number` `report_prefix` `excel_day_counter` `lookup` `lookup_join`.
- Unfilled placeholders stay visible as `{{name}}`: search the report for `{{`.
- Warnings are in **Details…**; after editing any file press **⟳ Refresh** (F5).

## 13. What to give an AI to get precise help

Use **Copy guide + my profile** in the guide window. It adds, after this guide:

1. the **check results** of your profile (the warnings you see in **Details…**),
2. your **mapping.json** (the folder paths of the `config` block are hidden),
3. the **placeholders found in your Word template**,
4. the **Excel sheet names and headers** (only the headers and row counts, never the cell values),

and a short instruction for the AI. Then describe the problem in your own words, for example:
"The `{{k_value}}` placeholder stays empty in the report" or "Add a field for the column `lot`".
