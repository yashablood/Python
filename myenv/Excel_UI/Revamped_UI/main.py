import tkinter as tk
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
        load_scripts()  # Load scripts after file selection

def load_scripts():
    # You can initialize or load data here from the scripts if necessary
    dashboard_script.load_data(file_path)
    data_script.load_data(file_path)
    error_tracker_script.load_data(file_path)
    otif_script.load_data(file_path)
    production_script.load_data(file_path)
    recognitions_script.load_data(file_path)

def submit_data():
    print("Submit button clicked")

    # Collect data from the entry fields
    data = {}
    for i, entry in enumerate(entry_fields):
        label = labels[i]
        data[label] = entry.get()  # Store the input data

    # Call update functions from each loaded module
    for module in sheet_modules:
        try:
            module.update_sheet(file_path, data)  # Call the update function for each module
        except AttributeError:
            print(f"No update function in {module.__name__}")    

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
