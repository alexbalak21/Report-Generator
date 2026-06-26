import os
import json
import datetime


class ReportStateManager:
    def __init__(self, state_path):
        self.state_path = state_path

    def load(self):
        if not os.path.exists(self.state_path):
            return {"last_report_number": None}

        with open(self.state_path, "r", encoding="utf-8") as f:
            return json.load(f)

    def save(self, state):
        with open(self.state_path, "w", encoding="utf-8") as f:
            json.dump(state, f, indent=2)

    def generate_report_number(self):
        state = self.load()
        last = state.get("last_report_number")

        today = datetime.date.today()
        prefix = today.strftime("%y%m%d")  # e.g. 260611

        if last and last.startswith(prefix):
            last_counter = int(last.split("-")[1])
            new_counter = last_counter + 1
        else:
            new_counter = 1

        new_number = f"{prefix}-{new_counter:02d}"

        state["last_report_number"] = new_number
        self.save(state)

        return new_number
