import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
from utils.excel_handler import load_workbook, save_workbook
import json
import os

CONFIG_FILE = "config.json"  # File to store the last file path


def save_last_file_path(file_path):
    """Save the last selected file path to a configuration file."""
    with open(CONFIG_FILE, "w") as f:
        json.dump({"last_file": file_path}, f)


def load_last_file_path():
    """Load the last selected file path from the configuration file."""
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, "r") as f:
            data = json.load(f)
            return data.get("last_file")
    return None


class DataEntryApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Professional Data Entry")
        self.geometry("800x600")
        self.resizable(True, True)

        self.workbook = None
        self.file_path = None
        self.sheet_mapping = {}
        self.fields = {}  # Holds StringVar instances for each input field
        self.field_to_sheet_mapping = {}  # Maps fields to sheets and cell locations

        # Attempt to auto-load the last file
        self.auto_load_last_file()

        # Initialize UI
        self.create_ui()

    def auto_load_last_file(self):
        """Auto-load the last selected file if it exists."""
        last_file = load_last_file_path()
        if last_file and os.path.exists(last_file):
            try:
                self.workbook = load_workbook(last_file)
                self.file_path = last_file
                self.sheet_mapping = {name: self.workbook[name] for name in self.workbook.sheetnames}
                print(f"Auto-loaded file: {last_file}")
            except Exception as e:
                print(f"Failed to auto-load file: {e}")
                tk.messagebox.showerror("Error", f"Failed to auto-load file: {e}")

    def create_ui(self):
        """Create the main UI."""
        # File selection button
        file_button = ttk.Button(self, text="Open Excel File", command=self.load_excel_file)
        file_button.pack(pady=10)

        # Notebook for tabs
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill="both", expand=True, padx=10, pady=10)

        # Tabs
        self.sheet1_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.sheet1_frame, text="Sheet 1")

        self.recognition_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.recognition_frame, text="Recognition Entry")

        # Define field mappings
        self.field_to_sheet_mapping = {
            # Sheet 1 fields
            "Days without Incident": ("Dashboard Rev 2", (4, 2)),
            "Haz ID's": ("Data", (3, 3)),
            "Safety Gemba Walk": ("Data", (4, 3)),
            "7S (Zone 26)": ("Data", (5, 3)),
            "7S (Zone 51)": ("Data", (6, 3)),

            # Recognition fields
            "First Name": ("Recognitions", (2, 1)),  # Starts at row 2
            "Last Name": ("Recognitions", (2, 2)),
            "Recognition": ("Recognitions", (2, 3)),
            "Date": ("Recognitions", (2, 4)),
        }

        # Initialize fields
        self.fields = {field: tk.StringVar() for field in self.field_to_sheet_mapping.keys()}

        # Add fields to tabs
        self.add_fields(self.sheet1_frame, ["Days without Incident", "Haz ID's", "Safety Gemba Walk", "7S (Zone 26)", "7S (Zone 51)"])
        self.add_fields(self.recognition_frame, ["First Name", "Last Name", "Recognition", "Date"])

        # Add save button
        save_button = ttk.Button(self, text="Save Data", command=self.save_data)
        save_button.pack(pady=10)

    def add_fields(self, frame, field_names):
        """Add labeled input fields for a given set of field names."""
        for idx, field_name in enumerate(field_names):
            ttk.Label(frame, text=field_name).grid(row=idx, column=0, padx=10, pady=5, sticky="w")
            ttk.Entry(frame, textvariable=self.fields[field_name]).grid(row=idx, column=1, padx=10, pady=5, sticky="ew")
            frame.columnconfigure(1, weight=1)

    def load_excel_file(self):
        """Load the Excel file and remember its path."""
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            try:
                self.workbook = load_workbook(file_path)
                self.file_path = file_path  # Store the selected file path
                save_last_file_path(file_path)  # Save the file path for future use
                self.sheet_mapping = {name: self.workbook[name] for name in self.workbook.sheetnames}
                print(f"Loaded file: {file_path}")
                print(f"Available sheets: {', '.join(self.sheet_mapping.keys())}")
            except Exception as e:
                print(f"Error loading Excel file: {e}")
                tk.messagebox.showerror("Error", f"Failed to load Excel file: {e}")

    def save_data(self):
        """Save all entered data to their corresponding Excel sheets."""
        if not self.workbook:
            tk.messagebox.showerror("Error", "No workbook loaded.")
            return

        try:
            for field_name, (sheet_name, (row, col)) in self.field_to_sheet_mapping.items():
                value = self.fields[field_name].get()  # Get user input
                if sheet_name in self.sheet_mapping:
                    sheet = self.sheet_mapping[sheet_name]
                    sheet.cell(row=row, column=col).value = value
                    print(f"Field '{field_name}' -> Sheet '{sheet_name}', Cell ({row},{col}): '{value}'")
                else:
                    print(f"Sheet '{sheet_name}' not found in the workbook.")

            # Save changes to the original file
            if self.file_path:
                save_workbook(self.workbook, self.file_path)
                print(f"Workbook saved to {self.file_path}")
                tk.messagebox.showinfo("Success", f"Data saved to {self.file_path}!")
            else:
                tk.messagebox.showerror("Error", "File path not set.")
        except Exception as e:
            print(f"Error saving data: {e}")
            tk.messagebox.showerror("Error", f"Failed to save data: {e}")

    def save_and_close(self):
        """Save the workbook and close the app."""
        if self.workbook:
            save_workbook(self.workbook, self.file_path)
            print("Workbook saved!")
        self.quit()


if __name__ == "__main__":
    app = DataEntryApp()
    app.mainloop()
