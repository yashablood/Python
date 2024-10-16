import tkinter as tk
from tkinter import filedialog

# Function to open a file dialog and select an Excel file
def select_file():
    file_path = filedialog.askopenfilename(title="Select Excel file", filetypes=[("Excel files", "*.xlsx")])
    if file_path:
        file_label.config(text=f"Selected File: {file_path}")
    else:
        file_label.config(text="No file selected")

# Function to handle submission (for now, this just prints data)
def submit_data():
    print("Submit button pressed")

# Setting up the main window
root = tk.Tk()
root.title("Excel Data Manipulation")
root.geometry("400x300")

# Top frame for the file selection button
top_frame = tk.Frame(root)
top_frame.pack(pady=10)

file_button = tk.Button(top_frame, text="Select Excel File", command=select_file)
file_button.pack()

file_label = tk.Label(top_frame, text="No file selected")
file_label.pack()

# Middle frame for the entry fields and labels
middle_frame = tk.Frame(root)
middle_frame.pack(pady=20)

# Labels on the left, Entry fields on the right (adjust as needed for number of fields)
labels = ["Days without Incident", "Haz ID's", "Safety Gemba Walk", "7S (Zone 26)", "7S (Zone 51)", "Errors", "PCD Returns",
        "Jobs on Hold", "Productivity", "OTIF %", "Huddles", "Truck Fill %", "Recognitions", "MC Compliance",
        "Cost Savings", "Rever's", "Project's", "Days Without Incident", "Haz ID's", "Safety Gemba Walk", "7S (Zone 26)", "7S (Zone 51)", "Errors", "PCD Returns", "Jobs on hold", "Productivity", "Otif", "Huddles", "Truck Fill", "Recognitions", "Master Control Compliance", "Cost Savings", " Rever's", "Projects"]  # Adjust based on the fields you want
entries = []

for label_text in labels:
    row_frame = tk.Frame(middle_frame)
    row_frame.pack(fill='x', pady=5)

    label = tk.Label(row_frame, text=label_text, width=15, anchor='w')
    label.pack(side='left')

    entry = tk.Entry(row_frame)
    entry.pack(side='right', fill='x', expand=True)
    entries.append(entry)

# Bottom frame for the submit button
bottom_frame = tk.Frame(root)
bottom_frame.pack(pady=20)

submit_button = tk.Button(bottom_frame, text="Submit", command=submit_data)
submit_button.pack()

# Start the Tkinter event loop
root.mainloop()
