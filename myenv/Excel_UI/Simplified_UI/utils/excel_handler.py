import openpyxl


def load_workbook(file_path):
    """Load an Excel workbook."""
    return openpyxl.load_workbook(file_path)


def save_workbook(workbook, file_path):
    """Save the Excel workbook."""
    workbook.save(file_path)


def get_sheet(workbook, sheet_name):
    """Retrieve a specific sheet by name."""
    return workbook[sheet_name]


def write_to_cell(sheet, row, col, value):
    """Write a value to a specific cell."""
    sheet.cell(row=row, column=col, value=value)


def read_cell(sheet, row, col):
    """Read a value from a specific cell."""
    return sheet.cell(row=row, column=col).value
