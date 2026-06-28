import os
import json
import datetime

from app.repository.rapport_repository import get_last_report_number, set_last_report_number


class ReportStateManager:
    def __init__(self, state_path: str):
        self._legacy_path = state_path
        self._migrate_if_needed()

    def _migrate_if_needed(self):
        if not os.path.exists(self._legacy_path):
            return
        if get_last_report_number() is not None:
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

    @staticmethod
    def get_today_prefix() -> str:
        return datetime.date.today().strftime("%y%m%d")

    @staticmethod
    def get_current_counter() -> int:
        """Return the current daily counter (0 if none yet today)."""
        last = get_last_report_number()
        prefix = datetime.date.today().strftime("%y%m%d")
        if last and last.startswith(prefix + "-"):
            try:
                return int(last.split("-")[1])
            except (IndexError, ValueError):
                pass
        return 0

    @staticmethod
    def set_counter(value: int) -> None:
        """Manually override the daily counter (persists to DB)."""
        prefix = datetime.date.today().strftime("%y%m%d")
        set_last_report_number(f"{prefix}-{value:02d}")

    def generate_report_number(self) -> str:
        """
        Generate the next daily counter string: yymmdd-NN
        (no prefix — the prefix comes from the mapping config and is
        applied by the 'numéro_rapport' computed field in the mapping).
        """
        last = get_last_report_number()
        prefix = self.get_today_prefix()

        if last and last.startswith(prefix + "-"):
            try:
                new_counter = int(last.split("-")[1]) + 1
            except (IndexError, ValueError):
                new_counter = 1
        else:
            new_counter = 1

        new_number = f"{prefix}-{new_counter:02d}"
        set_last_report_number(new_number)
        return new_number