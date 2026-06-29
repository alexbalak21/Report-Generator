"""
Handles report generation logic triggered from the GUI.
"""
import os
import datetime
import traceback

from app.core.report_generator import ReportGenerator
from app.core.mapping_loader import MappingLoader
from app.core.excel_reader import ExcelReader
from app.core.state_manager import ReportStateManager
from app.repository.rapport_repository import save_report


def _peek_report_number(state_mgr: ReportStateManager) -> str:
    """
    Return what the next report number WILL be, without incrementing the counter.
    Used only for the filename preview shown in the Save As dialog.
    """
    current = state_mgr.get_current_counter()
    prefix  = state_mgr.get_today_prefix()
    return f"{prefix}-{current + 1:02d}"


def resolve_output_path(mapping_path: str, excel_path: str,
                        docx_path: str, row_number: int) -> tuple[str, str]:
    """
    Compute the suggested (directory, filename) for the Save As dialog.
    Priority for output dir: SQLite last_output_dir > JSON output_dir > cwd.
    Returns (initial_dir, initial_filename).
    """
    from app.repository.config_repository import config_get
    try:
        loader = MappingLoader(mapping_path)
        cfg    = loader.load_config()
        rules  = loader.load()
    except Exception:
        return os.getcwd(), f"report_row_{row_number}.docx"

    initial_dir   = config_get("last_output_dir") or cfg.get("output_dir") or os.getcwd()
    date_format   = cfg.get("date_format", "%d/%m/%Y")
    report_prefix = cfg.get("report_prefix", "")

    file_name_rule = rules.get("file_name")
    if not isinstance(file_name_rule, dict):
        return initial_dir, f"report_row_{row_number}.docx"

    try:
        gen = ReportGenerator(excel_path, docx_path, mapping_path)

        # Peek at the next report number WITHOUT consuming it
        peeked_number = _peek_report_number(gen.state_manager)

        # Build a fake operations map that returns the peeked number instead of
        # calling generate_report_number() (which would increment the counter)
        from app.core import processors

        def _peek_report_number_op(rule, row):
            return peeked_number

        operations = {
            "today":             processors.op_today,
            "uppercase":         processors.op_uppercase,
            "lowercase":         processors.op_lowercase,
            "format":            processors.op_format,
            "concat":            processors.op_concat,
            "report_number":     _peek_report_number_op,
            "report_prefix":     lambda rule, row: report_prefix,
            "excel_day_counter": lambda rule, row: processors.op_excel_day_counter(
                                     rule, row, excel, row_number),
            "lookup":            lambda rule, row: processors.op_lookup(rule, row, excel),
            "lookup_join":       lambda rule, row: processors.op_lookup_join(rule, row, excel),
        }

        excel = ExcelReader(excel_path)
        excel.load()
        raw_row  = excel.get_row_as_dict(row_number)
        row_data = {
            k: gen._normalize_field_value(v, date_format)
            for k, v in raw_row.items()
        }

        # Seed computed with known values
        computed = {"report_prefix": report_prefix}

        # Resolve all computed fields except file_name (single pass is enough here)
        for _ in range(len(rules) + 1):
            resolved_any = False
            for key, rule in rules.items():
                if not isinstance(rule, dict) or key == "file_name" or key in computed:
                    continue
                enriched = {**row_data, **computed}
                try:
                    val = gen.compute_value(rule, enriched, operations)
                    if val is not None:
                        computed[key] = str(val)
                        resolved_any = True
                except Exception:
                    pass
            if not resolved_any:
                break

        # Resolve file_name
        enriched = {**row_data, **computed}
        raw_name = gen.compute_value(file_name_rule, enriched, operations)
        if raw_name:
            return initial_dir, gen._sanitize_filename(raw_name)

    except Exception:
        traceback.print_exc()   # visible in the terminal — helps diagnose future issues

    return initial_dir, f"report_row_{row_number}.docx"


def generate_report(excel: str, docx: str, mapping: str,
                    row_number: int, output_path: str) -> str:
    """
    Run the report generator to the explicit output_path chosen by the user.
    Returns the final path.
    """
    generator = ReportGenerator(
        excel_path=excel,
        template_path=docx,
        mapping_path=mapping,
    )
    final_path = generator.generate(row_number=row_number, output_path=output_path)
    _log_report(excel, docx, mapping, row_number)
    return final_path


def _log_report(excel: str, docx: str, mapping: str, row_number: int) -> None:
    try:
        save_report(
            report_id=f"row_{row_number}_{datetime.datetime.now().strftime('%Y%m%d%H%M%S')}",
            created_at=datetime.datetime.now().isoformat(),
            excel_path=excel,
            template_path=docx,
            mapping_path=mapping,
            row_number=row_number,
            data={},
        )
    except Exception:
        pass
