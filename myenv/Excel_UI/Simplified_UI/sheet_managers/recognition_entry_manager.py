from utils.excel_handler import get_sheet, write_to_cell


class RecognitionEntryManager:
    def __init__(self, workbook):
        self.workbook = workbook
        self.sheet = get_sheet(workbook, "Recognition Entry")

    def add_recognition(self, recognition):
        """
        Add a new recognition entry to the sheet.
        :param recognition: Dictionary with keys `First Name`, `Last Name`, `Recognition`, and `Date`.
        """
        # Find the next empty row
        row = self.sheet.max_row + 1

        # Write recognition data to the sheet
        write_to_cell(self.sheet, row, 1, recognition["First Name"])
        write_to_cell(self.sheet, row, 2, recognition["Last Name"])
        write_to_cell(self.sheet, row, 3, recognition["Recognition"])
        write_to_cell(self.sheet, row, 4, recognition["Date"])
