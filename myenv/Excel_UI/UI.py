import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
import pandas as pd
import openpyxl 
import os
import matplotlib.pyplot as plt
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows

# Load the Excel file
file_path = 'Boxing Tier.xlsx'
print("File path:", os.path.abspath(file_path))


def adjust_column_widths(ws):
    for col in ws.iter_cols():
        max_length = 0
        for cell in col:
            try:
                if cell.value and len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        adjusted_width = (max_length + 2) * 1.2
        ws.column_dimensions[col[0].column_letter].width = adjusted_width


# Specify the worksheet name where the data resides
def process_file(file_path):
    try:    
        wb = load_workbook(filename=file_path)
        worksheet_names = wb.sheetnames

        for worksheet_name in worksheet_names:
        # Read data from the worksheet
            df = pd.read_excel(file_path, sheet_name=worksheet_name,)
            ws = wb[worksheet_name]
            adjust_column_widths(ws)

        # Save the changes to the Excel file
        wb.save(file_path)
        return "Processing complete!"
    except Exception as e:
        return f"An error occurred: {e}"

def select_file():
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

    update_column_names(file_path, worksheet_names[0])
    
def update_column_names(file_path, sheet_name):
    wb = load_workbook(filename=file_path)
    ws = wb[sheet_name]
    column_names = [cell.value for cell in ws[3]]
    column_name_combobox['values'] = column_names
    column_name_combobox.current(0)

def append_data_to_file():
    file_path = file_path_label.cget("text")
    if file_path == "No file selected":
        messagebox.showwarning("Warning", "Please select a file first.")
        return

    sheet_name = sheet_name_entry.get()
    column_name = column_name_entry.get()
    data = data_entry.get()

    if not sheet_name or not column_name or not data:
        messagebox.showwarning("Warning", "Please fill in all fields.")
        return

    result = append_data(file_path, sheet_name, column_name, data)
    result_text.insert(tk.END, result + "\n")

def add_new_sheet_or_column():
    file_path = file_path_label.cget("text")
    if file_path == "No file selected":
        messagebox.showwarning("Warning", "Please select a file first.")
        return

    new_sheet_name = new_sheet_name_entry.get()
    new_column_name = new_column_name_entry.get()

    if not new_sheet_name and not new_column_name:
        messagebox.showwarning("Warning", "Please provide a sheet name or a column name.")
        return

    wb = load_workbook(filename=file_path)

    if new_sheet_name:
        if new_sheet_name not in wb.sheetnames:
            wb.create_sheet(title=new_sheet_name)
        update_dropdowns(file_path)

    if new_column_name:
        sheet_name = sheet_name_combobox.get()
        ws = wb[sheet_name]
        column_names = [cell.value for cell in ws[3]]
        if new_column_name not in column_names:
            last_column_index = ws.max_column + 1
            ws.cell(row=3, column=last_column_index, value=new_column_name)
        update_column_names(file_path, sheet_name)

    wb.save(file_path)
    result_text.insert(tk.END, "New sheet or column added successfully!\n")

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

tk.Label(frame, text="Column Name:").pack()
column_name_combobox = ttk.Combobox(frame, state="readonly")
column_name_combobox.pack()

tk.Label(frame, text="Data to Append:").pack()
data_entry = tk.Entry(frame)
data_entry.pack()

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