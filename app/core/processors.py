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
    try:
        return fmt.format(**row_data)
    except Exception:
        return ""


def op_concat(rule, row_data):
    parts = rule.get("parts", [])
    return "".join(str(row_data.get(p, "")) for p in parts)
