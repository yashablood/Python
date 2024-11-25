import tkinter as tk
import ui_controller 
import pandas as pd  # Add this at the top of the script
from tkinter import filedialog
from tkcalendar import DateEntry
from scripts import (
    dashboard_script, 
    data_script, 
    error_tracker_script, 
    otif_script, 
    production_script, 
    recognitions_script
    )

# Define sheet_modules to store the imported script modules
sheet_modules = [
    dashboard_script,
    data_script,
    error_tracker_script,
    otif_script,
    production_script,
    recognitions_script,
]

def open_file():
    global file_path
    file_path = filedialog.askopenfilename(title="Select Excel File", filetypes=[("Excel files", "*.xlsx")])
    if file_path:
        print(f"Selected file: {file_path}")
        
        # Debugging check
        print("Attempting to read and display date formats from the sheet.")

        try:
            sheet_names = pd.ExcelFile(file_path).sheet_names
            print(f"Available sheets: {sheet_names}")
        except Exception as e:
            print(f"Error fetching sheet names: {e}")


        # Load and display date formats from the Excel sheet
        try:
            sheet_name = "Data"  # Replace with the correct sheet name
            print(f"Opening sheet: {sheet_name}")
            df = pd.read_excel(file_path, sheet_name=sheet_name)

            if 'Date' in df.columns:
                print("\nDates in Excel sheet (raw):")
                print(df['Date'].head())  # Display raw dates from the Excel sheet

                # Ensure the 'Date' column is interpreted as datetime objects
                df['Date'] = pd.to_datetime(df['Date'], errors='coerce')
                print("\nDates in Excel sheet (after conversion to datetime):")
                print(df['Date'].head())
            else:
                print("'Date' column not found in the sheet.")
        except Exception as e:
            print(f"Error reading Excel file: {e}")
        
        if 'Date' not in df.columns:
            print(f"'Date' column not found in {sheet_name}. Available columns: {df.columns}")

        print("Preview of the loaded sheet:")
        print(df.head())

        # Debugging check
        print("Finished processing dates. Proceeding to load scripts.")

        # Load scripts after reading and debugging date formats
        load_scripts()
        
        open_file()
        print("open_file function completed.")

def load_scripts():
    # You can initialize or load data here from the scripts if necessary
    for module in sheet_modules:
        try:
            module.load_data(file_path)
            print(f"{module.__name__} loaded data successfully.")
        except AttributeError:
            print(f"{module.__name__} does not have a load_data function.")
        except Exception as e:
            print(f"Error loading {module.__name__}: {e}")

def submit_data():
   
    global file_path  # Ensure file_path is accessible here
    if not file_path:
        print("No file selected. Please select a file before submitting data.")
        return
    
    print("Submit button clicked")

       # Collect data from the entry fields
    data = {}
    for i, entry in enumerate(entry_fields):
        label = labels[i]
        data[label] = entry.get()  # Store the input data

    # Get the selected date from the DateEntry widget
    selected_date = date_entry.get_date().strftime("%Y-%m-%d") + " 00:00:00"
    print(f"Formatted selected date: {selected_date}")
    
    # Iterate through each module to update the appropriate sheet via ui_controller
    #sheet_names = ["Dashboard Rev 2", "Data", "Error Tracker", "OTIF", "Production", "Recognitions"]
    #for sheet_name in sheet_names:
        # Pass control to ui_controller, which will handle specific scripts and logic
        #ui_controller.update_sheet(sheet_name, file_path, data, selected_date)

# Call update functions from each loaded module
    for module in sheet_modules:
        try:
            if module == data_script:
                # Pass date only to data_script's update function
                module.update_data_sheet(file_path, data, selected_date)
            else:
                # Call update without date for other modules
                module.update_sheet(file_path, data)
        except AttributeError:
            print(f"No update function in {module.__name__}")
        except Exception as e:
            print(f"Error in {module.__name__}: {e}")

# Initialize the main window
root = tk.Tk()
root.title("Data Entry UI")
root.geometry("400x400")  # Set the window size

# Create a frame to hold all widgets
main_frame = tk.Frame(root)
main_frame.pack(fill="both", expand=True)

# Create a Canvas and a Scrollbar
canvas = tk.Canvas(main_frame)
scrollbar = tk.Scrollbar(main_frame, orient="vertical", command=canvas.yview)
scrollable_frame = tk.Frame(canvas)

# Configure the scrollable frame
scrollable_frame.bind(
    "<Configure>",
    lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
)

canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
canvas.configure(yscrollcommand=scrollbar.set)

# Add the scrollbar to the right side of the main frame
canvas.pack(side="left", fill="both", expand=True)
scrollbar.pack(side="right", fill="y")

# Button to select Excel file
file_button = tk.Button(scrollable_frame, text="Select Excel File", command=open_file)
file_button.grid(row=0, column=0, columnspan=2, pady=(0, 20))

# Calendar for date selection
date_label = tk.Label(scrollable_frame, text="Select Date:")
date_label.grid(row=1, column=0, sticky="e", padx=5, pady=5)
date_entry = DateEntry(scrollable_frame, width=12, background='darkblue', foreground='white', borderwidth=2)
date_entry.grid(row=1, column=1, padx=5, pady=5)

# Labels and Entry fields for data input
labels = [
    "Days without Incident", "Haz ID's", "Safety Gemba Walk", "7S (Zone 26)", "7S (Zone 51)", "Errors", "PCD Returns",
    "Jobs on Hold", "Productivity", "OTIF %", "Huddles", "Truck Fill %", "Recognitions", "MC Compliance",
    "Cost Savings", "Rever's", "Project's"
]

entry_fields = []

# Start adding labels and entry fields from row 2
for i, label in enumerate(labels):
    tk.Label(scrollable_frame, text=label).grid(row=i+2, column=0, sticky="e", padx=5, pady=5)  # Adjust row index
    entry = tk.Entry(scrollable_frame)
    entry.grid(row=i+2, column=1, padx=5, pady=5)  # Adjust row index
    entry_fields.append(entry)

# Submit button at the bottom
submit_button = tk.Button(scrollable_frame, text="Submit", command=submit_data)
submit_button.grid(row=len(labels) + 2, column=0, columnspan=2, pady=20)

# Enable scrolling with the mouse wheel
def on_mouse_wheel(event):
    canvas.yview_scroll(int(-1*(event.delta/120)), "units")

canvas.bind_all("<MouseWheel>", on_mouse_wheel)

# Run the application
root.mainloop()
