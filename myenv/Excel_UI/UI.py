import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
import pandas as pd
import openpyxl 
import os
import matplotlib.pyplot as plt
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.utils import get_column_letter

# Global variables
file_path = None
current_sheet_name = None
entries = {}

script_dir = os.path.dirname(__file__)

# Load the Excel file
file_path = 'Boxing Tier.xlsx'
print("File path:", os.path.abspath(file_path))

def adjust_column_widths(ws):
    for col in ws.iter_cols():
        max_length = 0
        column_letter = get_column_letter(col[0].column)
        for cell in col:
            try:
                # Skip merged cells
                if cell.is_merged_cell:
                    continue
                
                # Calculate the length of the cell value
                if cell.value and len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except Exception as e:
                print(f"Error processing cell {cell.coordinate}: {e}")

        adjusted_width = (max_length + 2) * 1.2
        ws.column_dimensions[column_letter].width = adjusted_width


# Specify the worksheet name where the data resides
def process_file(file_path):
    try:    
        wb = load_workbook(filename=file_path)
        # Optionally save all the sheetnames to global variable for later use.
        worksheet_names = wb.sheetnames

        for worksheet_name in worksheet_names:
            # Read data from the worksheet
            df = pd.read_excel(file_path, sheet_name=worksheet_name,)
            ws = wb[worksheet_name]
            adjust_column_widths(ws)

            # Read required sheets
            Dashboard_Rev_2_df = pd.read_excel(file_path, sheet_name='Dashboard Rev 2')
            Data_df = pd.read_excel(file_path, sheet_name='Data')
            Recognitions_df = pd.read_excel(file_path, sheet_name='Recognitions')
            Error_Tracker_df = pd.read_excel(file_path, sheet_name='Error Tracker')
            Production_df = pd.read_excel(file_path, sheet_name='Production')
            OTIF_df = pd.read_excel(file_path, sheet_name='OTIF')

        # Save the changes to the Excel file
        wb.save(file_path)
        return "Processing complete!"
    except Exception as e:
        return f"An error occurred: {e}"

# Select a file and update the dropdowns
def select_file():
    global file_path
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if file_path:
        file_path_label.config(text=os.path.basename(file_path))
        result = process_file(file_path)
        update_dropdowns(file_path)
        result_text.insert(tk.END, result + "\n")

def update_dropdowns(file_path):
    wb = load_workbook(filename=file_path)
    worksheet_names = wb.sheetnames
    sheet_name_combobox['values'] = worksheet_names
    sheet_name_combobox.current(0)
    #update_column_names()
    update_sheet_window()

def update_sheet_window(event=None):
    global current_sheet_name
    sheet_name = sheet_name_combobox.get()
    current_sheet_name = sheet_name

    # Destroy previous widgets in the sheet window
    for widget in sheet_window_frame.winfo_children():
        widget.destroy()

    # Determine which sheet-specific function to call based on sheet_name
    if sheet_name == 'Dashboard Rev 2':
        create_dashboard_rev_2_window()
    elif sheet_name == 'Data':
        create_data_window()
    elif sheet_name == 'Recognitions':
        create_recognitions_window()
    elif sheet_name == 'Error Tracker':
        create_error_tracker_window()
    elif sheet_name == 'Production':
        create_production_window()
    elif sheet_name == 'OTIF':
        create_otif_window()
    else:
        # Handle unknown sheet name
        pass

def create_dashboard_rev_2_window():
    # Define specific widgets and layout for 'Dashboard Rev 2' sheet
    pass

def create_data_window():
    # Define specific widgets and layout for 'Data' sheet
    pass

def create_recognitions_window():
    # Define specific widgets and layout for 'Recognitions' sheet
    def update_column_names(event=None):
        try:
            global file_path
            wb = load_workbook(filename=file_path)
            sheet_name = sheet_name_combobox.get()
            ws = wb[sheet_name]
            column_names = []

                # Ensure that we are only getting column names from the first row
            for cell in ws[1]:
            #if cell.value is not None:  # Skip if the cell is empty
                column_names.append(cell.value)
                create_dynamic_form(column_names)
        except Exception as e:
            result_text.insert(tk.END, f"An error occurred: {e}\n")                
        pass

def create_error_tracker_window():
    # Define specific widgets and layout for 'Error Tracker' sheet
    pass

def create_production_window():
    # Define specific widgets and layout for 'Production' sheet
    pass

def create_otif_window():
    # Define specific widgets and layout for 'OTIF' sheet
    pass

    
#def update_column_names(event=None):
    #try:
        #global file_path
        #wb = load_workbook(filename=file_path)
        #sheet_name = sheet_name_combobox.get()
        #ws = wb[sheet_name]
        #column_names = []

        # Ensure that we are only getting column names from the first row
        #for cell in ws[1]:
            #if cell.value is not None:  # Skip if the cell is empty
                #column_names.append(cell.value)

        #create_dynamic_form(column_names)
    #except Exception as e:
        #result_text.insert(tk.END, f"An error occurred: {e}\n")

def create_dynamic_form(column_names):
    for widget in dynamic_frame.winfo_children():
        widget.destroy()

    #global entries
    #entries = {}

    for column in column_names:
        if column is None:
            continue  # Skip if column name is None


        label = tk.Label(dynamic_frame, text=column)
        label.pack()

        if column.lower().startswith("date"):
            entry = tk.Entry(dynamic_frame)
        elif column.lower().startswith("long text"):
            entry = tk.Text(dynamic_frame, height=4, width=30)
        else:
            entry = tk.Entry(dynamic_frame)
        entry.pack()
        
        entries[column] = entry

def append_data_to_file():
    file_path = file_path_label.cget("text")
    if file_path == "No file selected":
        messagebox.showwarning("Warning", "Please select a file first.")
        return
    

    sheet_name = sheet_name_combobox.get()
    #column_name = column_name_combobox.get()
    #sheet_name = sheet_name_entry.get()
    #column_name = column_name_entry.get()
    #data = data_entry.get()
    data = {}
    for column, entry in entries.items():
        if isinstance(entry, tk.Text):
            data[column] = entry.get("1.0", tk.END).strip()
        else:
            data[column] = entry.get().strip()

    if not sheet_name or not data:
        messagebox.showwarning("Warning", "Please fill in all fields.")
        return

    result = append_data(file_path, sheet_name, data)
    result_text.insert(tk.END, result + "\n")

def append_data(file_path, sheet_name, data):
    try:
        wb = load_workbook(file_path)
        ws = wb[sheet_name]

        new_row = ws.max_row + 1
        for column, value in data.items():
            col_index = ws[1].index(column) + 1
            ws.cell(row=new_row, column=col_index, value=value)

        wb.save(file_path)
        return "Data appended successfully!"
    except Exception as e:
        return f"An error occurred: {e}"


# Create the main window
root = tk.Tk()
root.title("Sample Data Processor")

# Create a frame for the file selection
frame = tk.Frame(root)
frame.pack(padx=10, pady=10)

# Add a label and button for file selection
file_path_label = tk.Label(frame, text="No file selected")
file_path_label.pack(anchor="w")

select_file_button = tk.Button(frame, text="Select File", command=select_file)
select_file_button.pack()

# Add entry fields for appending data
entry_frame = tk.Frame(root)
entry_frame.pack(padx=10, pady=10)

tk.Label(frame, text="Sheet Name:").pack()
sheet_name_combobox = ttk.Combobox(frame, state="readonly")
sheet_name_combobox.pack()
sheet_name_combobox.bind("<<ComboboxSelected>>", update_sheet_window)

# Frame for sheet-specific widgets
sheet_window_frame = tk.Frame(root)
sheet_window_frame.pack(padx=10, pady=10)

# Frame for dynamic entries
dynamic_frame = tk.Frame(root)
dynamic_frame.pack(padx=10, pady=10)

#tk.Label(frame, text="Column Name:").pack()
#column_name_combobox = ttk.Combobox(frame, state="readonly")
#column_name_combobox.pack()

#tk.Label(frame, text="Data to Append:").pack()
#data_entry = tk.Entry(frame)
#data_entry.pack()

append_data_button = tk.Button(frame, text="Append Data", command=append_data_to_file)
append_data_button.pack()

    #Add these on another optional page
#tk.Label(frame, text="New Sheet Name:").pack()
#new_sheet_name_entry = tk.Entry(frame)
#new_sheet_name_entry.pack()

#tk.Label(frame, text="New Column Name:").pack()
#new_column_name_entry = tk.Entry(frame)
#new_column_name_entry.pack()

#add_new_button = tk.Button(frame, text="Add Sheet/Column", command=add_new_sheet_or_column)
#add_new_button.pack()

# Add a text widget for showing results
result_text = tk.Text(root, height=10, width=80)
result_text.pack(padx=10, pady=10)

# Start the GUI event loop
root.mainloop()