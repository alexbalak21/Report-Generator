import os
import sys


def _base_dir() -> str:
    """Root directory: PyInstaller bundle root when frozen, project root in dev."""
    if getattr(sys, "frozen", False):
        # sys.executable is dist/report-generator/report-generator.exe
        return os.path.dirname(sys.executable)
    # In dev: this file is app/utils/paths.py → go up two levels to project root
    return os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))


def get_data_path(filename: str) -> str:
    return os.path.join(_base_dir(), "data", filename)


def get_resource_path(filename: str) -> str:
    """For files bundled at the root level (e.g. icon.ico)."""
    return os.path.join(_base_dir(), filename)
