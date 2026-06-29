import openpyxl


class ExcelReader:
    def __init__(self, filepath):
        self.filepath = filepath
        self.workbook = None
        # Cache: { sheet_name: [row_dict, ...] }
        self._sheets: dict[str, list[dict]] = {}

    def load(self):
        """Load the workbook and pre-cache all sheets."""
        self.workbook = openpyxl.load_workbook(self.filepath, data_only=True)
        self._sheets = {}
        for name in self.workbook.sheetnames:
            self._sheets[name] = self._read_sheet(self.workbook[name])

    # ------------------------------------------------------------------
    # Internal
    # ------------------------------------------------------------------

    @staticmethod
    def _normalize_header(value) -> str:
        if value is None:
            return ""
        return str(value).strip()

    def _read_sheet(self, sheet) -> list[dict]:
        """Return all data rows of a sheet as a list of dicts keyed by header."""
        rows = list(sheet.iter_rows(values_only=True))
        if not rows:
            return []
        headers = [self._normalize_header(h) for h in rows[0]]
        result = []
        for row in rows[1:]:
            result.append({headers[i]: row[i] for i in range(len(headers))})
        return result

    def _ensure_loaded(self):
        if self.workbook is None:
            self.load()

    # ------------------------------------------------------------------
    # Public API — single active sheet (backwards compatible)
    # ------------------------------------------------------------------

    @property
    def sheet(self):
        """Return the active sheet (first sheet), for backwards compatibility."""
        self._ensure_loaded()
        return self.workbook.active

    @property
    def worksheet(self):
        """Alias — used by processors that need direct cell access."""
        return self.sheet

    def get_columns(self) -> list[str]:
        """Return normalized column names from the first (active) sheet."""
        self._ensure_loaded()
        active = self.workbook.active.title
        rows = self._sheets.get(active, [])
        return list(rows[0].keys()) if rows else []

    def get_row_as_dict(self, row_number: int) -> dict:
        """
        Return a row from the active sheet as a dict.
        row_number is 1-based Excel row (row 1 = header, data starts at 2).
        """
        self._ensure_loaded()
        active = self.workbook.active.title
        rows = self._sheets.get(active, [])
        data_index = row_number - 2  # row 2 → index 0
        if data_index < 0 or data_index >= len(rows):
            raise IndexError(f"Row {row_number} out of range in sheet '{active}'.")
        return rows[data_index]

    # ------------------------------------------------------------------
    # Public API — multi-sheet
    # ------------------------------------------------------------------

    def get_sheet(self, sheet_name: str) -> list[dict]:
        """
        Return all data rows of a named sheet as a list of dicts.
        Raises KeyError if the sheet does not exist.
        """
        self._ensure_loaded()
        if sheet_name not in self._sheets:
            available = ", ".join(self._sheets.keys())
            raise KeyError(
                f"Sheet '{sheet_name}' not found. Available sheets: {available}"
            )
        return self._sheets[sheet_name]

    def find_row(self, sheet_name: str, key_column: str, key_value) -> dict | None:
        """
        Return the first row in sheet_name where key_column == key_value.
        Returns None if no match is found.
        Comparison is done as strings (stripped) to handle mixed int/str Excel values.
        """
        rows = self.get_sheet(sheet_name)
        target = str(key_value).strip() if key_value is not None else ""
        for row in rows:
            cell = row.get(key_column)
            if str(cell).strip() if cell is not None else "" == target:
                return row
        return None