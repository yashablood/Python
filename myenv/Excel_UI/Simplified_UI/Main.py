import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
from utils.excel_handler import load_workbook, save_workbook
from sheet_managers.recognition_entry_manager import RecognitionEntryManager
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
        self.recognition_fields = {}  # Holds StringVar instances for recognition input fields
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

                # Check for Recognitions sheet
                if "Recognitions" in self.sheet_mapping:
                    self.recognition_manager = RecognitionEntryManager(self.workbook)
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
            "Days without Incident": ("Dashboard Rev 2", (4, 2)),
            "Haz ID's": ("Data", (3, 3)),
            "Safety Gemba Walk": ("Data", (4, 3)),
            "7S (Zone 26)": ("Data", (5, 3)),
            "7S (Zone 51)": ("Data", (6, 3)),
        }

        # Define recognition fields separately
        self.recognition_fields = {
            "First Name": tk.StringVar(),
            "Last Name": tk.StringVar(),
            "Recognition": tk.StringVar(),
            "Date": tk.StringVar(),
        }

        # Initialize fields
        self.fields = {field: tk.StringVar() for field in self.field_to_sheet_mapping.keys()}

        # Add a LabelFrame for Sheet 1 fields
        sheet1_group = ttk.LabelFrame(self.sheet1_frame, text="Data Fields", padding=(10, 10))
        sheet1_group.pack(fill="both", expand=True, padx=10, pady=10)

        # Add fields to the LabelFrame
        for idx, (label, var) in enumerate(self.fields.items()):
            ttk.Label(sheet1_group, text=label).grid(row=idx, column=0, padx=10, pady=5, sticky="w")
            ttk.Entry(sheet1_group, textvariable=var).grid(row=idx, column=1, padx=10, pady=5, sticky="ew")
            sheet1_group.columnconfigure(1, weight=1)

        # Add fields to the Recognition Entry tab
        self.add_fields(self.recognition_frame, self.recognition_fields)

        # Add save button
        save_button = ttk.Button(self, text="Save Data", command=self.save_data)
        save_button.pack(pady=10)

        # Bind the Enter key to the Save Data button
        save_button.bind("<Return>", lambda event: self.save_data())


    def add_fields(self, frame, fields):
        """Add labeled input fields for a given set of fields."""
        for idx, (field_name, var) in enumerate(fields.items()):
            ttk.Label(frame, text=field_name).grid(row=idx, column=0, padx=10, pady=5, sticky="w")
            ttk.Entry(frame, textvariable=var).grid(row=idx, column=1, padx=10, pady=5, sticky="ew")
            frame.columnconfigure(1, weight=1)

    def load_excel_file(self):
        """Load the Excel file and remember its path."""
        print(f"Loading workbook from {self.file_path}...")
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            try:
                self.workbook = load_workbook(file_path)
                self.file_path = file_path  # Store the selected file path
                save_last_file_path(file_path)  # Save the file path for future use
                self.sheet_mapping = {name: self.workbook[name] for name in self.workbook.sheetnames}

                # Check for Recognitions sheet
                if "Recognitions" in self.sheet_mapping:
                    print("Initializing RecognitionEntryManager...")
                    self.recognition_manager = RecognitionEntryManager(self.workbook)
                    print("RecognitionEntryManager initialized successfully.")
                else:
                    print("Recognitions sheet is missing.")

                print(f"Loaded file: {file_path}")
            except Exception as e:
                print(f"Error loading Excel file: {e}")
                tk.messagebox.showerror("Error", f"Failed to load Excel file: {e}")

    def save_data(self):
        """Save data based on the active tab."""
        if not self.workbook:
            tk.messagebox.showerror("Error", "No workbook loaded.")
            return

        try:
            # Get the index of the active tab
            active_tab = self.notebook.index(self.notebook.select())

            if active_tab == 0:  # Sheet 1 tab
                for field_name, (sheet_name, (row, col)) in self.field_to_sheet_mapping.items():
                    value = self.fields[field_name].get()  # Get user input
                    if sheet_name in self.sheet_mapping:
                        sheet = self.sheet_mapping[sheet_name]
                        sheet.cell(row=row, column=col).value = value
                        print(f"Field '{field_name}' -> Sheet '{sheet_name}', Cell ({row},{col}): '{value}'")
                    else:
                        print(f"Sheet '{sheet_name}' not found in the workbook.")

                if self.file_path:
                    try:
                        print(f"Attempting to save workbook to: {self.file_path}")
                        save_workbook(self.workbook, self.file_path)
                        print(f"Workbook saved to {self.file_path}")
                        tk.messagebox.showinfo("Success", f"Data saved to {self.file_path}!")

                        # Reset focus to the first entry in Sheet 1
                        first_field_entry = self.sheet1_frame.grid_slaves(row=0, column=1)[0]
                        first_field_entry.focus_set()
                    except PermissionError:
                        tk.messagebox.showerror(
                            "Error",
                            f"The file {self.file_path} is open in another program. Please close it and try again.",
                        )
                    except Exception as e:
                        tk.messagebox.showerror("Error", f"An unexpected error occurred: {e}")
                else:
                    tk.messagebox.showerror("Error", "File path not set.")



            elif active_tab == 1:  # Recognition Entry tab
                if hasattr(self, "recognition_manager") and self.recognition_manager:
                    recognition_data = {k: v.get() for k, v in self.recognition_fields.items()}
                    self.recognition_manager.add_recognition(recognition_data, self.file_path)
                    print(f"Recognition data saved: {recognition_data}")
                    tk.messagebox.showinfo("Success", "Recognition data saved successfully!")

                    # Reset focus to the first entry in Recognition Entry
                    first_field_entry = self.recognition_frame.grid_slaves(row=0, column=1)[0]
                    first_field_entry.focus_set()
                else:
                    tk.messagebox.showerror("Error", "Recognition manager not initialized. Please load a valid Excel file.")

            else:
                # General save for other tabs
                for field_name, (sheet_name, (row, col)) in self.field_to_sheet_mapping.items():
                    value = self.fields[field_name].get()  # Get user input
                    if sheet_name in self.sheet_mapping:
                        sheet = self.sheet_mapping[sheet_name]
                        sheet.cell(row=row, column=col).value = value
                        print(f"Field '{field_name}' -> Sheet '{sheet_name}', Cell ({row},{col}): '{value}'")
                    else:
                        print(f"Sheet '{sheet_name}' not found in the workbook.")

                if self.file_path:
                    try:
                        print(f"Attempting to save workbook to: {self.file_path}")
                        save_workbook(self.workbook, self.file_path)
                        print(f"Workbook saved to {self.file_path}")
                        tk.messagebox.showinfo("Success", f"Data saved to {self.file_path}!")
                    except PermissionError:
                        tk.messagebox.showerror(
                            "Error",
                            f"The file {self.file_path} is open in another program. Please close it and try again.",
                        )
                    except Exception as e:
                        tk.messagebox.showerror("Error", f"An unexpected error occurred: {e}")
                else:
                    tk.messagebox.showerror("Error", "File path not set.")

        except Exception as e:
            print(f"Error saving data: {e}")
            tk.messagebox.showerror("Error", f"Failed to save data: {e}")


if __name__ == "__main__":
    app = DataEntryApp()
    app.mainloop()