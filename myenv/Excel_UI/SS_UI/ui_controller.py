from scripts import dashboard_script
from scripts import data_script
from scripts import error_tracker_script
from scripts import otif_script 
from scripts import production_script 
from scripts import recognitions_script

def update_sheet(sheet_name, file_path, data, date=None):
    if sheet_name == 'Data':
        data_script.update_data_sheet(file_path, data, date)
        print(f"Data successfully updated in 'Data' sheet for date {date}.")
    elif sheet_name == 'Recognitions':
        recognitions_script.append_recognitions(file_path, data)
        print("Data successfully appended in 'Recognitions' sheet.")
    elif sheet_name == 'Error Tracker':
        error_tracker_script.append_error(file_path, data)
        print("Data successfully appended in 'Error Tracker' sheet.")
    elif sheet_name == 'Dashboard Rev 2':
        dashboard_script.overwrite_dashboard(file_path, data)
        print("Data successfully overwritten in 'Dashboard Rev 2'.")
    elif sheet_name == 'Production':
        production_script.handle_production(file_path, data)
        print("Data successfully updated in 'Production' sheet.")
    elif sheet_name == 'OTIF':
        otif_script.handle_otif(file_path, data)
        print("Data successfully updated in 'OTIF' sheet.")
    else:
        print(f"No handling defined for {sheet_name}.")
