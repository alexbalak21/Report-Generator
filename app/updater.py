import json
import threading
import webbrowser
import tkinter as tk
from tkinter import messagebox
from urllib.request import urlopen
from urllib.error import URLError

from app import __version__

GITHUB_API_URL = "https://api.github.com/repos/alexbalak21/Report-Generator/releases/latest"


def _parse_version(tag: str) -> tuple:
    """Convert 'v1.2.3' or '1.2.3' to (1, 2, 3) for comparison."""
    return tuple(int(x) for x in tag.lstrip("v").split("."))


def _fetch_latest() -> dict | None:
    """Fetch latest release info from GitHub. Returns None on any error."""
    try:
        with urlopen(GITHUB_API_URL, timeout=5) as response:
            return json.loads(response.read().decode())
    except (URLError, Exception):
        return None


def _show_update_dialog(latest_version: str, download_url: str) -> None:
    """Show a Tkinter dialog prompting the user to update."""
    root = tk.Tk()
    root.withdraw()

    message = (
        f"A new version is available: v{latest_version}\n"
        f"You are running: v{__version__}\n\n"
        f"Would you like to download the update?"
    )

    if messagebox.askyesno("Update Available", message, icon="info"):
        webbrowser.open(download_url)

    root.destroy()


def check_for_updates() -> None:
    """
    Check for updates in a background thread so the app starts normally.
    Shows a dialog only if a newer version is found.
    """
    def _check():
        data = _fetch_latest()
        if not data:
            return

        latest_tag = data.get("tag_name", "")
        if not latest_tag:
            return

        try:
            if _parse_version(latest_tag) <= _parse_version(__version__):
                return  # Already up to date
        except ValueError:
            return

        # Find the setup .exe asset, fallback to releases page
        assets = data.get("assets", [])
        download_url = next(
            (a["browser_download_url"] for a in assets if a["name"].endswith(".exe")),
            data.get("html_url", "https://github.com/alexbalak21/Report-Generator/releases/latest"),
        )

        _show_update_dialog(latest_tag.lstrip("v"), download_url)

    threading.Thread(target=_check, daemon=True).start()
