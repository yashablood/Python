import openpyxl
import logging

# Configure logging
logging.basicConfig(
    filename="error_log.txt",
    level=logging.ERROR,
    format="%(asctime)s - %(levelname)s - %(message)s",
)


def load_workbook(file_path):
    """Load an Excel workbook."""
    try:
        return openpyxl.load_workbook(file_path)
    except FileNotFoundError:
        logging.error(f"File not found: {file_path}")
        raise FileNotFoundError(f"The file {file_path} does not exist.")
    except Exception as e:
        logging.error(f"Failed to load workbook: {file_path} - {e}")
        raise


def save_workbook(workbook, file_path):
    """Save the workbook to the specified file path."""
    try:
        print(f"Attempting to save workbook to: {file_path}")
        workbook.save(file_path)
        print(f"Workbook successfully saved to: {file_path}")
    except Exception as e:
        print(f"Error saving workbook: {e}")
        raise


def get_sheet(workbook, sheet_name):
    """Retrieve a specific sheet by name."""
    try:
        # Debug: Log available sheets
        available_sheets = list(workbook.sheetnames)
        print(f"get_sheet called. Available sheets: {available_sheets}")

        return workbook[sheet_name]
    except KeyError:
        logging.error(f"Sheet not found: {sheet_name}. Available sheets: {available_sheets}")
        print(f"Sheet not found: {sheet_name}. Available sheets: {available_sheets}")
        raise KeyError(f"The sheet {sheet_name} does not exist in the workbook.")



def write_to_cell(sheet, row, col, value):
    """Write a value to a specific cell in the sheet."""
    try:
        sheet.cell(row=row, column=col).value = value
        print(f"Value '{value}' written to Row {row}, Column {col} in Sheet '{sheet.title}'")
    except Exception as e:
        print(f"Error writing to cell ({row}, {col}): {e}")
        raise


def read_cell(sheet, row, col):
    """Read a value from a specific cell."""
    try:
        return sheet.cell(row=row, column=col).value
    except Exception as e:
        logging.error(f"Failed to read cell: row={row}, col={col} - {e}")
        raise