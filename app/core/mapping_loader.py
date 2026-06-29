import json


class MappingLoader:
    def __init__(self, mapping_path):
        self.mapping_path = mapping_path

    def load(self) -> dict:
        """Return only the field mapping rules (excludes the 'config' key)."""
        raw = self._load_raw()
        return {k: v for k, v in raw.items() if k != "config"}

    def load_config(self) -> dict:
        """Return the 'config' block, or an empty dict if absent."""
        return self._load_raw().get("config", {})

    def update_config(self, **kwargs) -> None:
        """
        Update specific keys inside the 'config' block and write back to disk.
        Example: update_config(data_file="C:/Reports/table.xlsx")
        """
        raw = self._load_raw()
        raw.setdefault("config", {}).update(kwargs)
        with open(self.mapping_path, "w", encoding="utf-8") as f:
            json.dump(raw, f, indent=2, ensure_ascii=False)

    def update_file_name_field(self, **kwargs) -> None:
        """
        Update specific keys inside the 'file_name' rule and write back to disk.
        Example: update_file_name_field(name="NOVOCIB Rapport d'essai")
        """
        raw = self._load_raw()
        if "file_name" not in raw or not isinstance(raw["file_name"], dict):
            raw["file_name"] = {"operation": "format", "format": "{name} {numero_rapport}.docx"}
        raw["file_name"].update(kwargs)
        with open(self.mapping_path, "w", encoding="utf-8") as f:
            json.dump(raw, f, indent=2, ensure_ascii=False)

    def load_file_name_field(self) -> dict:
        """Return the 'file_name' rule dict, or empty dict if absent."""
        return self._load_raw().get("file_name", {})

    def _load_raw(self) -> dict:
        with open(self.mapping_path, "r", encoding="utf-8") as f:
            return json.load(f)