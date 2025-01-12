import openpyxl
import logging
import os
import json
from datetime import datetime

CONFIG_FILE = os.path.join(os.path.dirname(__file__), "config.json")
logging.info(f"Resolved CONFIG_FILE path: {CONFIG_FILE}")

# Configure logging
logging.basicConfig(
    filename="error_log.txt",
    level=logging.ERROR,
    format="%(asctime)s - %(levelname)s - %(message)s",
)


def extend_date_row(sheet, start_column):
    """Extend the date row in the Excel sheet for missing dates up until today."""
    try:
        today = datetime.now().date()

        # Find the last populated date in the date row
        last_date = None
        last_column = start_column - 1
        for col in range(start_column, sheet.max_column + 1):
            cell_value = sheet.cell(row=1, column=col).value
            if cell_value:
                if isinstance(cell_value, datetime):  # Handle proper datetime values
                    parsed_date = cell_value.date()
                elif isinstance(cell_value, str):  # Handle string dates
                    try:
                        parsed_date = datetime.strptime(cell_value, "%d-%b").date()
                    except ValueError:
                        parsed_date = None

                if parsed_date:
                    last_date = parsed_date
                    last_column = col

        # If no date is found, initialize with January 1 of the current year
        if not last_date:
            last_date = datetime(today.year, 1, 1).date()

        # Start extending dates from the day after the last date
        next_date = last_date + timedelta(days=1)
        current_column = last_column + 1

        # Add all missing dates until the current date
        while next_date <= today:
            # Check if the current date already exists in the date row
            is_duplicate = any(
                sheet.cell(row=1, column=col).value == next_date
                for col in range(start_column, sheet.max_column + 1)
            )
            if not is_duplicate:
                cell = sheet.cell(row=1, column=current_column)
                cell.value = next_date
                cell.number_format = "dd-mmm"  # Ensure consistent formatting
                logging.info(f"Added date {next_date.strftime('%d-%b')} to column {current_column}.")
                current_column += 1
            next_date += timedelta(days=1)

    except Exception as e:
        logging.error(f"Error extending date row: {e}")


def save_days_without_incident_data(counter, last_date):
    """Save the Days without Incident data to a JSON file."""
    try:
        config = load_config()
        config["counter"] = counter
        config["last_date"] = last_date.strftime("%Y-%m-%d")
        save_config(config)
        logging.info(f"Days without Incident updated: {config}")
    except Exception as e:
        logging.error(f"Error saving Days without Incident data: {e}")


def load_days_without_incident_data():
    """Load the Days without Incident data from a JSON file."""
    if not CONFIG_FILE:
        raise ValueError("CONFIG_FILE is not set.")
    if not os.path.exists(CONFIG_FILE):
        return {"counter": 0, "last_date": datetime.now().strftime("%Y-%m-%d")}
    try:
        with open(CONFIG_FILE, "r") as f:
            return json.load(f)
    except Exception as e:
        logging.error(f"Error loading Days without Incident data: {e}")
        raise


def find_or_add_date_column(sheet, selected_date, start_column=3):
    """Find or add the column for the selected date."""
    try:
        for col in range(start_column, sheet.max_column + 1):
            cell_value = sheet.cell(row=1, column=col).value
            if cell_value and isinstance(cell_value, datetime) and cell_value.date() == selected_date:
                return col
        # Add the date if not found
        new_col = sheet.max_column + 1
        write_to_cell(sheet, 1, new_col, selected_date)
        return new_col
    except Exception as e:
        logging.error(f"Error finding or adding date column: {e}")
        raise


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
        logging.error(
            f"Sheet not found: {sheet_name}. Available sheets: {available_sheets}")
        print(
            f"Sheet not found: {sheet_name}. Available sheets: {available_sheets}")
        raise KeyError(
            f"The sheet {sheet_name} does not exist in the workbook.")


def write_to_cell(sheet, row, col, value):
    """Write a value to a specific cell in the sheet."""
    try:
        sheet.cell(row=row, column=col).value = value
        print(
            f"Value '{value}' written to Row {row}, Column {col} in Sheet '{sheet.title}'")
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


def find_or_add_date_column(sheet, selected_date, start_column=3):
    """
    Find the column for the selected date in the first row of the sheet.
    If the date is not found, add it to the next available column.
    """
    try:
        # Check each column in the first row starting from start_column
        for col in range(start_column, sheet.max_column + 1):
            cell_value = sheet.cell(row=1, column=col).value
            if cell_value:
                # Handle proper datetime and string-based dates
                if isinstance(cell_value, datetime) and cell_value.date() == selected_date:
                    return col
                elif isinstance(cell_value, str):
                    try:
                        parsed_date = datetime.strptime(cell_value, "%d-%b").date()
                        if parsed_date == selected_date:
                            return col
                    except ValueError:
                        continue  # Ignore invalid date strings

        # Add the date to the next available column if not found
        new_column = sheet.max_column + 1
        cell = sheet.cell(row=1, column=new_column)
        cell.value = selected_date
        cell.number_format = "dd-mmm"  # Format for consistency
        logging.info(f"Added new date: {selected_date.strftime('%d-%b')} at column {new_column}.")
        return new_column

    except Exception as e:
        logging.error(f"Error in find_or_add_date_column: {e}")
        raise


def load_config():
    """Load configuration data from the JSON file."""
    try:
        if not os.path.exists(CONFIG_FILE):
            logging.warning(f"Config file not found. Creating a new one at {CONFIG_FILE}.")
            return {}
        with open(CONFIG_FILE, "r") as f:
            return json.load(f)
    except Exception as e:
        logging.error(f"Error loading config: {e}")
        return {}


def save_config(data):
    """Save configuration data to the JSON file."""
    try:
        with open(CONFIG_FILE, "w") as f:
            json.dump(data, f, indent=4)  # Save in a readable format
        logging.info(f"Config saved to {CONFIG_FILE}")
    except Exception as e:
        logging.error(f"Error saving config: {e}")


def calculate_truck_fill_percentage(value):
    """
    Validate and calculate the Truck Fill % based on the maximum value of 26.
    :param value: The entered value as a number (float or int).
    :return: The calculated percentage as a string with a '%' symbol.
    :raises ValueError: If the value is not a number or is outside the valid range.
    """
    try:
        value = float(value)  # Ensure the value is numeric
        if 0 <= value <= 26:
            percentage = (value / 26) * 100  # Calculate percentage
            return f"{percentage:.2f}%"  # Format with the '%' symbol
        else:
            raise ValueError("Value must be between 0 and 26.")
    except ValueError:
        raise ValueError(
            "Invalid input. Please enter a numeric value between 0 and 26.")
