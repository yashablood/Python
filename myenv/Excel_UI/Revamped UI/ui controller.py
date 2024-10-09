import data_script
import recognitions_script
import error_tracker_script
# Import other sheet scripts as needed

def update_sheet(sheet_name, file_path, data, date=None):
    if sheet_name == 'Data':
        data_script.update_data_sheet(file_path, data, date)
    elif sheet_name == 'Recognitions':
        recognitions_script.append_recognitions(file_path, data)
    elif sheet_name == 'Error Tracker':
        error_tracker_script.append_error(file_path, data)
    elif sheet_name == 'Dashboard Rev 2':
        dashboard_script.overwrite_dashboard(file_path, data)
    # Add other sheet handlers as needed

# Example call
if __name__ == "__main__":
    # This is an example of calling the function based on user input.
    file_path = "path_to_your_file.xlsx"
    data = {'Truck Fill %': 90}  # Example data
    date = '2024-10-10'

    update_sheet('Data', file_path, data, date)
