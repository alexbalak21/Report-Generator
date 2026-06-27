import os
import json
import datetime

from app.repository.rapport_repository import get_last_report_number, set_last_report_number


class ReportStateManager:
    def __init__(self, state_path: str):
        """
        state_path: legacy path to report_state.json — kept only for one-time migration.
        All state is now stored in the SQLite report_state table.
        """
        self._legacy_path = state_path
        self._migrate_if_needed()

    def _migrate_if_needed(self):
        """If report_state.json exists and SQLite has no entry yet, import it then delete it."""
        if not os.path.exists(self._legacy_path):
            return
        if get_last_report_number() is not None:
            # Already migrated — just remove the stale file
            try:
                os.remove(self._legacy_path)
            except OSError:
                pass
            return
        try:
            with open(self._legacy_path, "r", encoding="utf-8") as f:
                data = json.load(f)
            last = data.get("last_report_number")
            if last:
                set_last_report_number(last)
            os.remove(self._legacy_path)
        except Exception:
            pass

    def generate_report_number(self) -> str:
        last = get_last_report_number()

        today = datetime.date.today()
        prefix = today.strftime("%y%m%d")  # e.g. 260627

        if last and last.startswith(prefix):
            new_counter = int(last.split("-")[1]) + 1
        else:
            new_counter = 1

        new_number = f"{prefix}-{new_counter:02d}"
        set_last_report_number(new_number)
        return new_number
