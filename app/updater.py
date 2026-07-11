import json
import os
import subprocess
import tempfile
import threading
import tkinter as tk
from tkinter import ttk, messagebox
from urllib.request import urlopen, Request
from urllib.error import URLError

from app import __version__
from app.repository.config_repository import config_get, config_set

GITHUB_API_URL = "https://api.github.com/repos/alexbalak21/Report-Generator/releases/latest"
IGNORED_VERSION_KEY = "ignored_update_version"


def _parse_version(tag: str) -> tuple:
    return tuple(int(x) for x in tag.lstrip("v").split("."))


def _fetch_latest() -> dict | None:
    try:
        req = Request(GITHUB_API_URL, headers={"User-Agent": "ReportGenerator"})
        with urlopen(req, timeout=5) as response:
            return json.loads(response.read().decode())
    except (URLError, Exception):
        return None


class UpdateDialog(tk.Toplevel):
    """Dialog that downloads the installer and shows a progress bar."""

    def __init__(self, parent, latest_version: str, download_url: str):
        super().__init__(parent)
        self.title("Update Available")
        self.resizable(False, False)
        self.grab_set()  # modal

        self._download_url = download_url
        self._latest_version = latest_version

        self._build_ui(latest_version)
        self._center(parent)

    def _build_ui(self, latest_version: str):
        tk.Label(
            self,
            text="A new version is available!",
            font=("Arial", 13, "bold"),
        ).pack(padx=24, pady=(20, 4))

        tk.Label(
            self,
            text=f"v{__version__}  →  v{latest_version}",
            font=("Arial", 11),
            fg="#555",
        ).pack(padx=24, pady=4)

        tk.Label(
            self,
            text="The update will be downloaded and installed automatically.",
            wraplength=320,
        ).pack(padx=24, pady=4)

        self._status = tk.Label(self, text="", fg="#333")
        self._status.pack(padx=24, pady=4)

        self._progress = ttk.Progressbar(self, length=320, mode="determinate")
        self._progress.pack(padx=24, pady=4)

        # Buttons row
        btn_frame = tk.Frame(self)
        btn_frame.pack(pady=(8, 4))

        self._install_btn = tk.Button(
            btn_frame,
            text="Download & Install",
            width=18,
            bg="#0078D4",
            fg="white",
            font=("Arial", 10, "bold"),
            relief="flat",
            command=self._start_download,
        )
        self._install_btn.pack(side="left", padx=6)

        tk.Button(
            btn_frame,
            text="Later",
            width=10,
            relief="flat",
            command=self.destroy,
        ).pack(side="left", padx=6)

        # Ignore this version link-style button
        tk.Button(
            self,
            text="Ignore this version",
            font=("Arial", 8),
            fg="#888",
            relief="flat",
            bd=0,
            cursor="hand2",
            command=self._ignore_version,
        ).pack(pady=(0, 16))

    def _center(self, parent):
        self.update_idletasks()
        x = parent.winfo_rootx() + (parent.winfo_width() - self.winfo_width()) // 2
        y = parent.winfo_rooty() + (parent.winfo_height() - self.winfo_height()) // 2
        self.geometry(f"+{x}+{y}")

    def _ignore_version(self):
        """Store the ignored version in DB so this version is never shown again."""
        config_set(IGNORED_VERSION_KEY, self._latest_version)
        self.destroy()

    def _start_download(self):
        self._install_btn.config(state="disabled", text="Downloading…")
        self._status.config(text="Starting download…")
        threading.Thread(target=self._download, daemon=True).start()

    def _download(self):
        try:
            req = Request(self._download_url, headers={"User-Agent": "ReportGenerator"})
            with urlopen(req, timeout=30) as response:
                total = int(response.headers.get("Content-Length", 0))
                downloaded = 0
                chunk_size = 8192

                tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".exe")
                tmp_path = tmp.name

                while True:
                    chunk = response.read(chunk_size)
                    if not chunk:
                        break
                    tmp.write(chunk)
                    downloaded += len(chunk)

                    if total:
                        pct = downloaded / total * 100
                        mb_done = downloaded / 1_048_576
                        mb_total = total / 1_048_576
                        self.after(0, self._update_progress, pct, mb_done, mb_total)

                tmp.close()

            self.after(0, self._launch_installer, tmp_path)

        except Exception as e:
            self.after(0, self._on_error, str(e))

    def _update_progress(self, pct: float, mb_done: float, mb_total: float):
        self._progress["value"] = pct
        self._status.config(text=f"Downloading… {mb_done:.1f} / {mb_total:.1f} MB")

    def _launch_installer(self, installer_path: str):
        self._status.config(text="Download complete — launching installer…")
        self._progress["value"] = 100
        subprocess.Popen([installer_path], shell=False)
        self.after(500, self._quit_app)

    def _quit_app(self):
        root = self.master
        self.destroy()
        root.quit()
        root.destroy()

    def _on_error(self, message: str):
        self._install_btn.config(state="normal", text="Download & Install")
        self._status.config(text="")
        messagebox.showerror(
            "Download Failed",
            f"Could not download the update:\n{message}\n\n"
            "Please try again or download manually from GitHub.",
            parent=self,
        )


def check_for_updates(parent_window: tk.Tk) -> None:
    """
    Check for updates in a background thread.
    If a newer version is found (and not ignored), show the UpdateDialog.
    """
    def _check():
        data = _fetch_latest()
        if not data:
            return

        latest_tag = data.get("tag_name", "")
        if not latest_tag:
            return

        latest_version = latest_tag.lstrip("v")

        try:
            if _parse_version(latest_tag) <= _parse_version(__version__):
                return  # Already up to date
        except ValueError:
            return

        # Check if this version was previously ignored
        ignored = config_get(IGNORED_VERSION_KEY)
        if ignored and _parse_version(ignored) >= _parse_version(latest_version):
            return  # User chose to ignore this version

        assets = data.get("assets", [])
        download_url = next(
            (a["browser_download_url"] for a in assets if a["name"].endswith(".exe")),
            None,
        )

        if not download_url:
            return

        parent_window.after(0, lambda: UpdateDialog(parent_window, latest_version, download_url))

    threading.Thread(target=_check, daemon=True).start()
