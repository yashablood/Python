import tkinter as tk
from tkinter import filedialog
import ui_controller

def select_file():
    # Open a file dialog to select the Excel file
    file_path = filedialog.askopenfilename(
        title="Select an Excel file",
        filetypes=[("Excel Files", "*.xlsx;*.xls")]
    )
    if file_path:
        print(f"Selected file: {file_path}")
        return file_path
    else:
        print("No file selected")
        return None

def update_data():
    # Get the entered data from the UI
    truck_fill = entry_truck_fill.get()
    days_without_incident = entry_days_without_incident.get()
    date = entry_date.get()

    # Select the Excel file to update
    file_path = select_file()

    if file_path:
        # Create the data dictionary to pass to the update function
        data = {'Truck Fill %': truck_fill, 'Days without Incident': days_without_incident}
        
        # Update the sheet using the data entered in the UI
        ui_controller.update_sheet('Data', file_path, data, date)

def create_ui():
    root = tk.Tk()
    root.title("Excel Sheet Updater")

    # Labels and entry fields
    tk.Label(root, text="Truck Fill %:").grid(row=0, column=0, padx=10, pady=5)
    global entry_truck_fill
    entry_truck_fill = tk.Entry(root)
    entry_truck_fill.grid(row=0, column=1, padx=10, pady=5)

    tk.Label(root, text="Days without Incident:").grid(row=1, column=0, padx=10, pady=5)
    global entry_days_without_incident
    entry_days_without_incident = tk.Entry(root)
    entry_days_without_incident.grid(row=1, column=1, padx=10, pady=5)

    tk.Label(root, text="Date (MM/DD/YYYY):").grid(row=2, column=0, padx=10, pady=5)
    global entry_date
    entry_date = tk.Entry(root)
    entry_date.grid(row=2, column=1, padx=10, pady=5)

    # Submit button
    update_button = tk.Button(root, text="Update Excel Sheet", command=update_data)
    update_button.grid(row=3, column=0, columnspan=2, pady=10)

    root.mainloop()

if __name__ == "__main__":
    create_ui()
