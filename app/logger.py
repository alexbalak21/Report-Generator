"""
Centralised logging for Report-Generator.

Log location
------------
Installed exe  : %APPDATA%\\ReportGenerator\\logs\\app_YYYY-MM-DD.log
Dev (py run.py): <project_root>/logs/app_YYYY-MM-DD.log

A new file is created each day; the last 30 files are kept automatically.

Usage
-----
    from app.logger import get_logger
    log = get_logger(__name__)

    log.info("Report generated: %s", path)
    log.warning("Template placeholder not found: %s", key)
    log.error("Failed to load workbook", exc_info=True)   # includes traceback
    log.debug("Row data: %s", row_dict)
"""

import logging
import logging.handlers
import os
import sys
from datetime import date

# --------------------------------------------------------------------------- #
#  Resolve log directory                                                        #
# --------------------------------------------------------------------------- #

def _log_dir() -> str:
    if getattr(sys, "frozen", False):
        # Installed exe → %APPDATA%\ReportGenerator\logs\
        base = os.path.join(os.environ.get("APPDATA", os.path.expanduser("~")),
                            "ReportGenerator")
    else:
        # Dev → <project_root>/logs/
        base = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

    log_dir = os.path.join(base, "logs")
    os.makedirs(log_dir, exist_ok=True)
    return log_dir


# --------------------------------------------------------------------------- #
#  Build the root logger once                                                   #
# --------------------------------------------------------------------------- #

_LOGGER_NAME = "report_generator"
_initialised  = False


def _initialise() -> None:
    global _initialised
    if _initialised:
        return
    _initialised = True

    root = logging.getLogger(_LOGGER_NAME)
    root.setLevel(logging.DEBUG)          # capture everything; handlers filter

    fmt = logging.Formatter(
        fmt     = "%(asctime)s  %(levelname)-8s  %(name)s  —  %(message)s",
        datefmt = "%Y-%m-%d %H:%M:%S",
    )

    # ── File handler: one file per day, keep 30 days ──────────────────────── #
    log_file = os.path.join(_log_dir(), f"app_{date.today():%Y-%m-%d}.log")
    fh = logging.handlers.TimedRotatingFileHandler(
        filename    = log_file,
        when        = "midnight",
        interval    = 1,
        backupCount = 30,
        encoding    = "utf-8",
        delay       = False,
    )
    fh.setLevel(logging.DEBUG)
    fh.setFormatter(fmt)
    root.addHandler(fh)

    # ── Console handler: INFO+ in dev, nothing when frozen ────────────────── #
    if not getattr(sys, "frozen", False):
        ch = logging.StreamHandler(sys.stdout)
        ch.setLevel(logging.INFO)
        ch.setFormatter(fmt)
        root.addHandler(ch)

    # ── Redirect bare print() / warnings to the log ───────────────────────── #
    logging.captureWarnings(True)

    root.info("=" * 70)
    root.info("Report-Generator starting  (frozen=%s)", getattr(sys, "frozen", False))
    root.info("Log file: %s", log_file)
    root.info("=" * 70)


def get_logger(name: str) -> logging.Logger:
    """
    Return a child logger.  Call this at module level:

        log = get_logger(__name__)
    """
    _initialise()
    # Strip leading 'app.' so names are shorter: 'core.report_generator' etc.
    short = name.removeprefix("app.")
    return logging.getLogger(f"{_LOGGER_NAME}.{short}")
