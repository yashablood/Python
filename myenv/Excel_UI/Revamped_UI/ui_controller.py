from scripts import data_script
from scripts import recognitions_script
from scripts import error_tracker_script
from scripts import production_script  # Importing the stub script
from scripts import otif_script        # Importing the stub script
from scripts import dashboard_script

def update_sheet(sheet_name, file_path, data, date=None):
    if sheet_name == 'Data':
        data_script.update_data_sheet(file_path, data, date)
        print("Modules imported successfully!")
    elif sheet_name == 'Recognitions':
        recognitions_script.append_recognitions(file_path, data)
        print("Modules imported successfully!")
    elif sheet_name == 'Error Tracker':
        error_tracker_script.append_error(file_path, data)
        print("Modules imported successfully!")
    elif sheet_name == 'Dashboard Rev 2':
        dashboard_script.overwrite_dashboard(file_path, data)
        print("Modules imported successfully!")
    elif sheet_name == 'Production':
        production_script.handle_production(file_path, data)  # Calls the stub
        print("Modules imported successfully!")
    elif sheet_name == 'OTIF':
        otif_script.handle_otif(file_path, data)  # Calls the stub
        print("Modules imported successfully!")
    else:
        print(f"No handling defined for {sheet_name}.")


# Example call
#if __name__ == "__main__":
    # This is an example of calling the function based on user input.
    #file_path = "path_to_your_file.xlsx"
    #data = {'Truck Fill %': 90}  # Example data
    #date = '2024-10-10'

    #update_sheet('Data', file_path, data, date)
