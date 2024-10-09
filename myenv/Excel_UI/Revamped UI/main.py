import ui_controller

if __name__ == "__main__":
    # In real scenario, this could be input from a form, command line, etc.
    sheet_name = "Data"
    file_path = "path_to_your_excel_file.xlsx"
    data = {
        'Truck Fill %': 90,
        'Days without Incident': 5,
    }
    date = '2024-10-10'
    
    ui_controller.update_sheet(sheet_name, file_path, data, date)
