"""
Handles saving and restoring GUI state (paths, line number) via SQLite.
"""
from app.core.mapping_loader import MappingLoader
from app.repository.config_repository import (
    config_get,
    config_set,
    mapping_add,
)

KEY_EXCEL        = "last_excel_path"
KEY_DOCX         = "last_docx_path"
KEY_MAPPING      = "mapping_path"
KEY_MAPPING_LEGACY = "last_mapping_path"
KEY_LINE         = "last_line_number"
KEY_OUTPUT_DIR   = "last_output_dir"


def restore_config() -> dict:
    """
    Load last-used state from SQLite.
    Returns a dict with keys: excel, docx, mapping, line, output_dir, mapping_cfg.
    """
    excel      = config_get(KEY_EXCEL) or ""
    docx       = config_get(KEY_DOCX) or ""
    mapping    = config_get(KEY_MAPPING) or config_get(KEY_MAPPING_LEGACY) or ""
    output_dir = config_get(KEY_OUTPUT_DIR) or ""

    last_line = config_get(KEY_LINE)
    line = (int(last_line) + 1) if (last_line and last_line.isdigit()) else 2

    mapping_cfg = load_mapping_config(mapping) if mapping else {}

    return {
        "excel":       excel,
        "docx":        docx,
        "mapping":     mapping,
        "line":        line,
        "output_dir":  output_dir,
        "mapping_cfg": mapping_cfg,
    }


def save_config(excel: str, docx: str, mapping: str, line: int) -> None:
    """Persist current UI state to SQLite after a successful generation."""
    config_set(KEY_EXCEL,   excel)
    config_set(KEY_DOCX,    docx)
    config_set(KEY_MAPPING, mapping)
    config_set(KEY_MAPPING_LEGACY, mapping)
    config_set(KEY_LINE,    str(line))
    mapping_add(mapping)


def save_mapping_path(path: str) -> None:
    """Persist mapping path immediately when the user selects one."""
    config_set(KEY_MAPPING, path)
    config_set(KEY_MAPPING_LEGACY, path)
    mapping_add(path)


def save_output_dir(directory: str) -> None:
    """Persist the output directory chosen in the Save As dialog."""
    config_set(KEY_OUTPUT_DIR, directory)


def load_mapping_config(mapping_path: str) -> dict:
    """
    Read the 'config' block from a mapping JSON.
    Returns an empty dict on any error.
    """
    try:
        return MappingLoader(mapping_path).load_config()
    except Exception:
        return {}