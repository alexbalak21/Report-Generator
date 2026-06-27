"""
Handles report generation logic triggered from the GUI.
"""
import os
import datetime

from app.core.report_generator import ReportGenerator
from app.repository.rapport_repository import save_report


def generate_report(excel: str, docx: str, mapping: str, row_number: int) -> str:
    """
    Run the report generator and return the output path.
    Raises on any error — caller is responsible for showing the error to the user.
    """
    output_path = os.path.join(os.getcwd(), f"report_row_{row_number}.docx")

    generator = ReportGenerator(
        excel_path=excel,
        template_path=docx,
        mapping_path=mapping,
    )
    generator.generate(row_number=row_number, output_path=output_path)

    _log_report(excel, docx, mapping, row_number)

    return output_path


def _log_report(excel: str, docx: str, mapping: str, row_number: int) -> None:
    """Append report metadata to SQLite — silently ignored on failure."""
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
