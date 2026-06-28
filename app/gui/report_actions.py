"""
Handles report generation logic triggered from the GUI.
"""
import os
import datetime

from app.core.report_generator import ReportGenerator
from app.core.mapping_loader import MappingLoader
from app.core.excel_reader import ExcelReader
from app.repository.rapport_repository import save_report


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

    initial_dir = config_get("last_output_dir") or cfg.get("output_dir") or os.getcwd()
    date_format = cfg.get("date_format", "%d/%m/%Y")

    file_name_rule = rules.get("file_name")
    if isinstance(file_name_rule, dict):
        try:
            gen        = ReportGenerator(excel_path, docx_path, mapping_path)
            excel      = ExcelReader(excel_path)
            excel.load()
            raw_row    = excel.get_row_as_dict(row_number)

            # Normalize with the correct date_format from config
            row_data = {
                k: gen._normalize_field_value(v, date_format)
                for k, v in raw_row.items()
            }

            report_prefix = cfg.get("report_prefix", "")

            # Build the full operations map (needed for compute_value)
            operations = gen._build_operations(excel, report_prefix)

            # Resolve all computed fields except file_name
            computed = {"report_prefix": report_prefix}
            for key, rule in rules.items():
                if not isinstance(rule, dict) or key == "file_name":
                    continue
                enriched = {**row_data, **computed}
                try:
                    val = gen.compute_value(rule, enriched, operations)
                    if val is not None:
                        computed[key] = str(val)
                except Exception:
                    pass

            # Resolve file_name with all computed values available
            enriched = {**row_data, **computed}
            raw_name = gen.compute_value(file_name_rule, enriched, operations)
            if raw_name:
                return initial_dir, gen._sanitize_filename(raw_name)
        except Exception:
            pass

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