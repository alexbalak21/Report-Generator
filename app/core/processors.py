import datetime


def op_today(rule, row_data):
    fmt = rule.get("format", "%Y-%m-%d")
    return datetime.date.today().strftime(fmt)


def op_uppercase(rule, row_data):
    field = rule.get("input")
    if field in row_data:
        return str(row_data[field]).upper()


def op_lowercase(rule, row_data):
    field = rule.get("input")
    if field in row_data:
        return str(row_data[field]).lower()


def op_format(rule, row_data):
    fmt = rule.get("format", "")
    # Merge rule's own fields (e.g. "name") into the context so {name} resolves
    context = {**rule, **row_data}
    try:
        return fmt.format(**context)
    except Exception:
        return ""


def op_concat(rule, row_data):
    parts = rule.get("parts", [])
    return "".join(str(row_data.get(p, "")) for p in parts)


# ── Lookup helpers ─────────────────────────────────────────────────────────────

def _find_row(excel_reader, sheet_name: str, key_column: str, key_value) -> dict | None:
    """Wrapper around ExcelReader.find_row with graceful error handling."""
    try:
        return excel_reader.find_row(sheet_name, key_column, key_value)
    except KeyError as e:
        raise LookupError(str(e))


def _normalize(value) -> str:
    return str(value).strip() if value is not None else ""


def _to_number(x):
    if x is None:
        raise ValueError("None is not a number")
    if isinstance(x, (int, float)):
        return float(x)
    s = str(x).strip()
    if s == "":
        raise ValueError("empty string")
    s = s.replace(",", ".")
    return float(s)


def _apply_numeric_operation(v, op):
    try:
        num = _to_number(v)
    except Exception:
        return v

    operand = op.get("value", 0)
    try:
        operand = float(operand)
    except Exception:
        operand = 0.0

    op_type = op.get("type")
    if op_type == "multiply":
        return num * operand
    if op_type == "divide":
        if operand == 0:
            return num
        return num / operand
    if op_type == "add":
        return num + operand
    if op_type == "subtract":
        return num - operand
    return num


def _apply_round(v, op):
    decimals = int(op.get("decimals", 0))
    try:
        num = _to_number(v)
        rounded = round(num, decimals)
        return int(rounded) if decimals == 0 else rounded
    except Exception:
        return v


def _apply_date_format(v, op):
    import datetime as _dt

    fmt = op.get("format", "%d/%m/%Y")
    if isinstance(v, (_dt.date, _dt.datetime)):
        return v.strftime(fmt)

    s = str(v).strip()
    if not s:
        return v

    parsed = None
    try:
        parsed = _dt.datetime.fromisoformat(s)
    except Exception:
        try:
            parsed = _dt.datetime.strptime(s, fmt)
        except Exception:
            for f in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%m/%d/%Y"):
                try:
                    parsed = _dt.datetime.strptime(s, f)
                    break
                except Exception:
                    parsed = None
    if parsed:
        return parsed.strftime(fmt)
    return v


def _apply_string_operation(v, op):
    op_type = op.get("type")
    if op_type == "suffix":
        return f"{v}{op.get('value','')}"
    if op_type == "prefix":
        return f"{op.get('value','')}{v}"
    if op_type == "upper":
        return str(v).upper()
    if op_type == "lower":
        return str(v).lower()
    if op_type == "strip":
        return str(v).strip()
    if op_type == "replace":
        find    = op.get("find", "")
        replace = op.get("replace", "")
        # Support \n as a literal newline in the JSON value
        replace = replace.replace("\\n", "\n")
        return str(v).replace(find, replace)
    return v


# ── op_lookup ─────────────────────────────────────────────────────────────────

def op_lookup(rule: dict, row_data: dict, excel_reader=None) -> str | None:
    """
    Simple single-sheet foreign-key lookup.

    rule keys:
      sheet  — sheet name to search in
      key    — column in the target sheet to match against
      match  — column in row_data that holds the foreign key value
      value  — column in the target sheet whose value to return
    """
    if excel_reader is None:
        return None

    sheet_name  = rule.get("sheet", "")
    key_column  = rule.get("key", "")
    match_field = rule.get("match", "")
    value_col   = rule.get("value", "")

    if not all([sheet_name, key_column, match_field, value_col]):
        return None

    foreign_key = row_data.get(match_field)
    if foreign_key is None:
        return None

    row = _find_row(excel_reader, sheet_name, key_column, foreign_key)
    if row is None:
        return None

    return _normalize(row.get(value_col))


# ── op_lookup_join ────────────────────────────────────────────────────────────

def op_lookup_join(rule: dict, row_data: dict, excel_reader=None) -> str | None:
    """
    Multi-step relational lookup (unlimited JOIN chain).

    rule keys:
      steps  — list of step dicts, each with:
                  sheet   sheet name to search in
                  key     column in the target sheet to match against
                  match   column in row_data (or "<previous>") with the FK value
                  value   column name or list of column names to return
      format — Python format string applied to the fields collected across all steps
               (uses the column names as keys, e.g. "{street}, {city}")

    "<previous>" in a step's match field uses the result returned by the previous step.
    When value is a list, all named columns are collected into the shared namespace
    and the last step's single-value result is also stored as "<previous>" for the
    next step (though multi-value steps are typically the last step).
    """
    if excel_reader is None:
        return None

    steps = rule.get("steps", [])
    fmt   = rule.get("format", "")

    if not steps:
        return None

    previous_value = None        # result carried between steps
    collected: dict[str, str] = {}  # all named fields gathered across steps

    for i, step in enumerate(steps):
        sheet_name  = step.get("sheet", "")
        key_column  = step.get("key", "")
        match_field = step.get("match", "")
        value_spec  = step.get("value")   # str or list[str]

        if not all([sheet_name, key_column, match_field, value_spec]):
            return None

        # Resolve the foreign key value
        if match_field == "<previous>":
            if previous_value is None:
                return None
            fk_value = previous_value
        else:
            fk_value = row_data.get(match_field)
            if fk_value is None:
                return None

        row = _find_row(excel_reader, sheet_name, key_column, fk_value)
        if row is None:
            return None

        # Collect requested columns
        if isinstance(value_spec, list):
            for col in value_spec:
                collected[col] = _normalize(row.get(col))
            # For chaining purposes, if there's a single-value next step it would
            # need an explicit match field — multi-value steps end the chain.
            previous_value = None
        else:
            # Single column — result becomes <previous> for the next step
            val = _normalize(row.get(value_spec))
            collected[value_spec] = val
            previous_value = val

    # Apply format string using all collected fields
    if fmt:
        try:
            return fmt.format(**collected)
        except KeyError as e:
            raise LookupError(
                f"lookup_join format references unknown field {e}. "
                f"Available: {list(collected.keys())}"
            )

    # No format string — return the last step's value only
    return previous_value if previous_value is not None else ""

# ── excel_day_counter ─────────────────────────────────────────────────────────

def op_excel_day_counter(rule: dict, row_data: dict, excel_reader=None, row_number: int = 0) -> str | None:
    """
    Counts how many times the date in `date_column` for the current row
    has appeared in that column up to and including the current row,
    then returns a report number in the format yymmdd-N.

    rule keys:
      date_column   — Excel column name containing the report date (e.g. "date rapport")
      sample_column — Excel column name containing the sample number (e.g. "numero echantillon")
      date_format   — strptime format of that column's string values (e.g. "%d/%m/%Y")
    """
    import datetime

    if excel_reader is None or row_number == 0:
        return None

    date_column   = rule.get("date_column", "date rapport")
    sample_column = rule.get("sample_column")
    date_fmt      = rule.get("date_format", "%d/%m/%Y")

    # Get the date of the current row
    current_date_raw = row_data.get(date_column, "")
    if not current_date_raw:
        return None

    try:
        current_date = datetime.datetime.strptime(str(current_date_raw).strip(), date_fmt).date()
    except ValueError:
        # Maybe it's already a date object stored as string from normalization
        try:
            current_date = datetime.datetime.strptime(str(current_date_raw).strip(), "%Y-%m-%d").date()
        except ValueError:
            return None

    day_str = current_date.strftime("%y%m%d")
    if sample_column:
        sample_value = row_data.get(sample_column, "")
        if sample_value is None or str(sample_value).strip() == "":
            return None
        return f"{day_str}-{str(sample_value).strip()}"

    # Walk all data rows (row 2 .. current row) and count occurrences of current_date
    counter = 0
    ws = excel_reader.worksheet  # openpyxl worksheet

    # Find the column index for date_column
    headers = {}
    for cell in ws[1]:
        if cell.value is not None:
            headers[str(cell.value).strip()] = cell.column

    if date_column not in headers:
        return None

    col_idx = headers[date_column]

    for r in range(2, row_number + 1):
        cell_val = ws.cell(row=r, column=col_idx).value
        if cell_val is None:
            continue
        try:
            if isinstance(cell_val, (datetime.date, datetime.datetime)):
                row_date = cell_val.date() if isinstance(cell_val, datetime.datetime) else cell_val
            else:
                row_date = datetime.datetime.strptime(str(cell_val).strip(), date_fmt).date()
            if row_date == current_date:
                counter += 1
        except ValueError:
            continue

    day_str = current_date.strftime("%y%m%d")
    return f"{day_str}-{counter}"


def apply_operations(value, operations: list | None, row_data: dict | None = None, excel_reader=None):
    """
    Apply a list of small transformation operations to `value`.

    Supported operation types:
      - formula: noop (workbook is loaded with data_only=True so value is already computed)
      - multiply/divide/add/subtract (use numeric coercion)
      - round (decimals int)
      - suffix/prefix (string concatenation)
      - upper/lower/strip

    Returns the transformed value (may be numeric or string depending on ops).
    """
    if not operations:
        return value

    v = value

    for op in operations:
        t = op.get("type")
        if t == "formula":
            continue

        if t in ("multiply", "divide", "add", "subtract"):
            v = _apply_numeric_operation(v, op)
            continue

        if t == "round":
            v = _apply_round(v, op)
            continue

        if t == "date_format":
            v = _apply_date_format(v, op)
            continue

        v = _apply_string_operation(v, op)

    return v