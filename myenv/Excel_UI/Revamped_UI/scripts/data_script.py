from openpyxl import load_workbook
from datetime import datetime

def load_data(file_path):
    # Load data from the specified Excel file
    pass

def update_data_sheet(file_path, data, date_str):
    print("Selected Data sheet")
    # Load the workbook and select the "Data" sheet
    wb = load_workbook(file_path)
    sheet = wb['Data']

    for col in range(3, sheet.max_column + 1):
        cell_value = sheet.cell(row=1, column=col).value
        print(f"Column {col}: {cell_value}")  # Debugging print

    # Convert the input date string to a datetime object
    selected_date = datetime.strptime(date_str, "%m/%d/%Y")

    # Search for the date in row 1 (from C1 onwards)
    for col in range(3, sheet.max_column + 1):  # Starts at column C (index 3)
        cell_value = sheet.cell(row=1, column=col).value

        # Check if the cell value matches the selected date (in date format)
        if isinstance(cell_value, datetime) and cell_value.date() == selected_date.date():
            print(f"Found date {selected_date.strftime('%m/%d/%Y')} in column {col}")
            
            # Now, update the data in the corresponding column
            for row, (key, value) in enumerate(data.items(), start=2):
                sheet.cell(row=row, column=col).value = value
            break
    else:
        raise ValueError(f"Date {date_str} not found in Data sheet.")

    # Save the workbook after updating
    wb.save(file_path)
