from utils.excel_handler import get_sheet, write_to_cell


class Sheet1Manager:
    def __init__(self, workbook):
        self.workbook = workbook
        self.sheet = get_sheet(workbook, "Sheet 1")

    def save_metrics(self, metrics):
        """
        Save metrics data to the sheet.
        :param metrics: Dictionary of field names to values.
        """
        # Example mapping of fields to rows/columns in the sheet
        field_mapping = {
            "Days without Incident": (2, 2),
            "Haz ID's": (3, 2),
            "Safety Gemba Walk": (4, 2),
            "7S (Zone 26)": (5, 2),
            "7S (Zone 51)": (6, 2),
        }

        for field, value in metrics.items():
            row, col = field_mapping[field]
            write_to_cell(self.sheet, row, col, value)
