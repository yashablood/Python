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
    """Save the Excel workbook."""
    try:
        workbook.save(file_path)
    except Exception as e:
        logging.error(f"Failed to save workbook: {file_path} - {e}")
        raise


def get_sheet(workbook, sheet_name):
    """Retrieve a specific sheet by name."""
    try:
        return workbook[sheet_name]
    except KeyError:
        logging.error(f"Sheet not found: {sheet_name}")
        raise KeyError(f"The sheet {sheet_name} does not exist in the workbook.")


def write_to_cell(sheet, row, col, value):
    """Write a value to a specific cell."""
    try:
        sheet.cell(row=row, column=col, value=value)
    except Exception as e:
        logging.error(f"Failed to write to cell: row={row}, col={col}, value={value} - {e}")
        raise


def read_cell(sheet, row, col):
    """Read a value from a specific cell."""
    try:
        return sheet.cell(row=row, column=col).value
    except Exception as e:
        logging.error(f"Failed to read cell: row={row}, col={col} - {e}")
        raise
