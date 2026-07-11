import os
import sys


def _exe_dir() -> str:
    """Directory containing the exe (or project root in dev).
    Used for files placed there by the installer (e.g. /data folder)."""
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))


def _meipass_dir() -> str:
    """PyInstaller's temp extraction folder.
    Used for files declared in run.spec datas[]."""
    return getattr(sys, "_MEIPASS", _exe_dir())


def get_data_path(filename: str) -> str:
    """Files from the /data folder — installed by Inno Setup next to the exe."""
    return os.path.join(_exe_dir(), "data", filename)


def get_resource_path(filename: str) -> str:
    """Files bundled via PyInstaller datas[] (e.g. icon.ico)."""
    return os.path.join(_meipass_dir(), filename)
