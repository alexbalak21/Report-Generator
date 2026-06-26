import openpyxl

class ExcelReader:
    def __init__(self, filepath):
        self.filepath = filepath
        self.workbook = None
        self.sheet = None

    def load(self):
        """Load the Excel file and first sheet."""
        self.workbook = openpyxl.load_workbook(self.filepath)
        self.sheet = self.workbook.active

    def get_columns(self):
        """Return column names from the first row."""
        if not self.sheet:
            self.load()
        return [cell.value for cell in self.sheet[1]]

    def get_row_as_dict(self, row_number):
        """
        Return a row as a dictionary:
        { column_name: value }
        """
        if not self.sheet:
            self.load()

        columns = self.get_columns()
        row = self.sheet[row_number]

        return {
            columns[i]: row[i].value
            for i in range(len(columns))
        }
