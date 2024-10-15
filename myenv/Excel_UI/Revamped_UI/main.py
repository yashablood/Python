import tkinter as tk
from tkinter import filedialog
import ui_controller

def select_file():
    # Create a new Tkinter window
    root = tk.Tk()
    root.withdraw()  # Hide the main window

    print("Select a file")
    # Open the file dialog and get the selected file path
    file_path = filedialog.askopenfilename(
        title="Select an Excel file",
        filetypes=[("Excel Files", "*.xlsx;*.xls")]
    )

    # Print or return the file path
    if file_path:
        print(f"Selected file: {file_path}")
        return file_path
    else:
        print("No file selected")
        return None

def main():
    print("Select a file 2")
    # Call the file selection function
    file_path = select_file()  # Get the file path
    if file_path:  # Proceed only if a file was selected
        
        # Define the sheet name and data you want to update
        sheet_name = "Data"  # Change this as needed
        data = {'Truck Fill %': 90, 'Days without Incident': 5}  # Example data
        date = "10/10/2024"  # Example date

        # Update the sheet with the selected file
        #ui_controller.update_sheet(sheet_name, file_path)

        # Example for Recognitions
        # recognition_data = {'Employee Name': 'John Doe', 'Recognition': 'Great job!'}
        # ui_controller.update_sheet('Recognitions', file_path, recognition_data)

        # Example for Error Tracker
        # error_data = {'Error Description': 'Sample error', 'Resolved': 'Yes'}
        # ui_controller.update_sheet('Error Tracker', file_path, error_data)

        # You can keep the lines for Production and OTIF as placeholders or remove them
        # ui_controller.update_sheet('Production', file_path, {'Data': 'Not implemented'})
        # ui_controller.update_sheet('OTIF', file_path, {'Data': 'Not implemented'})

if __name__ == "__main__":
    main()  # Call the main function to run the program
