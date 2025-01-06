import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from utils.excel_handler import load_workbook, save_workbook, calculate_truck_fill_percentage
from sheet_managers.recognition_entry_manager import RecognitionEntryManager
from tkcalendar import DateEntry
from datetime import datetime, timedelta
import json
import os
import logging

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

CONFIG_FILE = "config.json"  # File to store the last file path
DAYS_WITHOUT_INCIDENT_FILE = "days_without_incident.json"


def save_last_file_path(file_path):
    """Save the last selected file path to a configuration file."""
    try:
        with open(CONFIG_FILE, "w") as f:
            json.dump({"last_file": file_path}, f)
    except Exception as e:
        logging.error(f"Failed to save last file path: {e}")


def load_last_file_path():
    """Load the last selected file path from the configuration file."""
    try:
        if os.path.exists(CONFIG_FILE):
            with open(CONFIG_FILE, "r") as f:
                data = json.load(f)
                return data.get("last_file")
    except Exception as e:
        logging.error(f"Failed to load last file path: {e}")
    return None


def manage_days_without_incident(reset_toggle):
    """Manage the Days without Incident counter."""
    try:
        if not os.path.exists(DAYS_WITHOUT_INCIDENT_FILE):
            data = {"counter": 0, "last_date": datetime.now().strftime("%Y-%m-%d")}
        else:
            with open(DAYS_WITHOUT_INCIDENT_FILE, "r") as f:
                data = json.load(f)

        last_date = datetime.strptime(data["last_date"], "%Y-%m-%d").date()
        today = datetime.now().date()

        if reset_toggle:
            new_value = 0
            logging.info("Days without Incident reset to 0.")
        else:
            days_passed = (today - last_date).days
            new_value = data["counter"] + days_passed if days_passed > 0 else data["counter"]

        data["counter"] = new_value
        data["last_date"] = today.strftime("%Y-%m-%d")

        with open(DAYS_WITHOUT_INCIDENT_FILE, "w") as f:
            json.dump(data, f)

        logging.info(f"Days without Incident updated to {new_value}.")
        return new_value

    except Exception as e:
        logging.error(f"Error managing Days without Incident: {e}")
        return 0


def extend_date_row(sheet, start_column):
    """Extend the date row in the Excel sheet for missing dates up until today."""
    try:
        today = datetime.now().date()
        end_of_year = datetime(today.year, 12, 31).date()

        # Find the last populated date in the date row
        last_date = None
        last_column = start_column - 1
        for col in range(start_column, sheet.max_column + 1):
            cell_value = sheet.cell(row=1, column=col).value
            if cell_value:
                if isinstance(cell_value, datetime):  # Handle proper datetime values
                    parsed_date = cell_value.date()
                elif isinstance(cell_value, str):  # Handle string dates
                    try:
                        parsed_date = datetime.strptime(cell_value, "%d-%b").date()
                    except ValueError:
                        parsed_date = None

                if parsed_date:
                    last_date = parsed_date
                    last_column = col

        # If no date is found, initialize with January 1 of the current year
        if not last_date:
            last_date = datetime(today.year, 1, 1).date()

        # Start extending dates from the day after the last date
        next_date = last_date + timedelta(days=1)
        current_column = last_column + 1

        # Add all missing dates until the current date (don't move on to other tasks until done)
        while next_date <= today:
            cell = sheet.cell(row=1, column=current_column)
            cell.value = next_date  # Add as datetime object
            cell.number_format = "dd-mmm"  # Ensure consistent formatting
            logging.info(f"Added date {next_date.strftime('%d-%b')} to column {current_column}.")
            next_date += timedelta(days=1)
            current_column += 1

    except Exception as e:
        logging.error(f"Error extending date row: {e}")




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
        self.date_selection = tk.StringVar()

        self.auto_load_last_file()
        self.create_ui()

    def auto_load_last_file(self):
        """Auto-load the last selected file if it exists."""
        last_file = load_last_file_path()
        if last_file and os.path.exists(last_file):
            try:
                self.workbook = load_workbook(last_file)
                self.file_path = last_file
                self.sheet_mapping = {name: self.workbook[name] for name in self.workbook.sheetnames}

                if "Recognitions" in self.sheet_mapping:
                    self.recognition_manager = RecognitionEntryManager(self.workbook)

                logging.info(f"Auto-loaded file: {last_file}")
            except Exception as e:
                logging.error(f"Failed to auto-load file: {e}")
                messagebox.showerror("Error", f"Failed to auto-load file: {e}")

    def adjust_window_size(self):
        """Adjust the window size based on the content."""
        self.update_idletasks()
        width = self.notebook.winfo_reqwidth() + 20
        height = self.notebook.winfo_reqheight() + 20
        self.geometry(f"{width}x{height}")

    def create_ui(self):
        """Create the main UI."""
        file_button = ttk.Button(self, text="Open Excel File", command=self.load_excel_file)
        file_button.pack(pady=5)

        container = ttk.Frame(self)
        container.pack(fill="both", expand=True, padx=5, pady=5)

        canvas = tk.Canvas(container)
        scrollbar = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)

        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        def on_mouse_wheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

        canvas.bind_all("<MouseWheel>", on_mouse_wheel)

        self.notebook = ttk.Notebook(scrollable_frame)
        self.notebook.pack(fill="both", expand=True, padx=5, pady=5)

        self.sheet1_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.sheet1_frame, text="Sheet 1")

        self.recognition_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.recognition_frame, text="Recognition Entry")

        self.field_to_sheet_mapping = {
            "Days without Incident": ("Data", (2, 3)),
            "Haz ID's": ("Data", (3, 3)),
            "Safety Gemba Walk": ("Data", (4, 3)),
            "7S (Zone 26)": ("Data", (5, 3)),
            "7S (Zone 51)": ("Data", (6, 3)),
            "Errors": ("Data", (7, 3)),
            "PCD Returns": ("Data", (8, 3)),
            "Jobs on Hold": ("Data", (9, 3)),
            "Productivity": ("Data", (10, 3)),
            "OTIF": ("Data", (11, 3)),
            "Huddles": ("Data", (12, 3)),
            "Truck Fill %": ("Data", (13, 3)),
            "Recognitions": ("Data", (14, 3)),
            "MC Compliance": ("Data", (15, 3)),
            "Cost Savings": ("Data", (16, 3)),
            "Rever's": ("Data", (17, 3)),
            "Project's": ("Data", (18, 3)),
        }

        self.recognition_fields = {
            "First Name": tk.StringVar(),
            "Last Name": tk.StringVar(),
            "Recognition": tk.StringVar(),
            "Date": tk.StringVar(),
        }

        self.fields = {field: tk.StringVar() for field in self.field_to_sheet_mapping.keys()}

        sheet1_group = ttk.LabelFrame(self.sheet1_frame, text="Data Fields", padding=(10, 10))
        sheet1_group.pack(fill="both", expand=True, padx=5, pady=5)

        self.date_selection = tk.StringVar()
        date_picker = DateEntry(sheet1_group, textvariable=self.date_selection, width=20,
                                date_pattern="MM/dd/yyyy", background='darkblue', foreground='white', borderwidth=2)
        date_picker.grid(row=0, column=1, padx=10, pady=5, sticky="ew")
        ttk.Label(sheet1_group, text="Select Date:").grid(row=0, column=0, padx=10, pady=5, sticky="w")

        self.reset_days_toggle = tk.BooleanVar(value=False)
        reset_button = ttk.Checkbutton(sheet1_group, text="Reset Days", variable=self.reset_days_toggle)
        reset_button.grid(row=1, column=2, padx=10, pady=5, sticky="ew")

        # Populate the Days without Incident field on UI load
        days_without_incident = manage_days_without_incident(self.reset_days_toggle.get())
        self.fields["Days without Incident"].set(days_without_incident)

        for idx, (label, var) in enumerate(self.fields.items(), start=1):
            ttk.Label(sheet1_group, text=label).grid(row=idx, column=0, padx=10, pady=5, sticky="w")
            ttk.Entry(sheet1_group, textvariable=var).grid(row=idx, column=1, padx=10, pady=5, sticky="ew")

        self.add_fields(self.recognition_frame, self.recognition_fields)

        save_button = ttk.Button(scrollable_frame, text="Save Data", command=self.save_data)
        save_button.pack(pady=10)

        save_button.bind("<Return>", lambda event: self.save_data())

        self.adjust_window_size()

    def add_fields(self, frame, fields):
        for idx, (field_name, var) in enumerate(fields.items()):
            ttk.Label(frame, text=field_name).grid(row=idx, column=0, padx=10, pady=5, sticky="w")
            ttk.Entry(frame, textvariable=var).grid(row=idx, column=1, padx=10, pady=5, sticky="ew")

    def load_excel_file(self):
        """Load the Excel file and ensure the date row is updated."""
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            try:
                self.workbook = load_workbook(file_path)
                self.file_path = file_path
                save_last_file_path(file_path)
                self.sheet_mapping = {name: self.workbook[name] for name in self.workbook.sheetnames}

                if "Recognitions" in self.sheet_mapping:
                    self.recognition_manager = RecognitionEntryManager(self.workbook)

                # Extend the date row in the "Data" sheet
                data_sheet = self.sheet_mapping.get("Data")
                if data_sheet:
                    extend_date_row(data_sheet, start_column=3)

                logging.info(f"Loaded file: {file_path}")
            except Exception as e:
                logging.error(f"Error loading Excel file: {e}")
                messagebox.showerror("Error", f"Failed to load Excel file: {e}")


    def save_data(self):
        if not self.workbook:
            messagebox.showerror("Error", "No workbook loaded.")
            return

        try:
            active_tab = self.notebook.index(self.notebook.select())

            # First, check the date row in the "Data" sheet and add the full year's dates if needed
            data_sheet = self.sheet_mapping.get("Data")
            if data_sheet:
                # Check if the full year is present
                today = datetime.now().date()
                start_column = 3  # Dates start from column 3 (C)

                # Check if the current date already exists
                current_date_column = None
                for col in range(start_column, data_sheet.max_column + 1):
                    cell_value = data_sheet.cell(row=1, column=col).value
                    if cell_value and isinstance(cell_value, datetime):
                        if cell_value.date() == today:
                            current_date_column = col
                            break

                # If current date is not found, add the dates up until today
                if not current_date_column:
                    logging.info(f"Current date {today.strftime('%d-%b')} not found. Adding missing dates.")
                    extend_date_row(data_sheet, start_column)

            # Continue with the original logic after ensuring the date row is complete
            if active_tab == 0:  # Sheet 1 tab
                selected_date = self.date_selection.get()
                if not selected_date:
                    messagebox.showerror("Error", "Please select a date.")
                    return

                try:
                    selected_date = datetime.strptime(selected_date, "%m/%d/%Y").date()
                except ValueError as e:
                    logging.error(f"Invalid date selected: {e}")
                    messagebox.showerror("Error", f"Invalid date format: {selected_date}")
                    return

                # Ensure the selected date exists in the date row
                date_column = None
                for col in range(2, data_sheet.max_column + 1):
                    cell_value = data_sheet.cell(row=1, column=col).value
                    if cell_value and isinstance(cell_value, str):
                        try:
                            cell_date = datetime.strptime(cell_value, "%m/%d/%Y").date()
                            if cell_date == selected_date:
                                date_column = col
                                break
                        except ValueError:
                            continue

                # If the date is not found, add it to the next available column
                if not date_column:
                    logging.info(f"Date {selected_date} not found. Adding it to the date row.")
                    date_column = data_sheet.max_column + 1
                    data_sheet.cell(row=1, column=date_column, value=selected_date.strftime("%m/%d/%Y"))

                # Save data to the appropriate column
                for field_name, (sheet_name, (row, _)) in self.field_to_sheet_mapping.items():
                    if sheet_name == "Data":
                        value = self.fields[field_name].get()
                        data_sheet.cell(row=row, column=date_column, value=value)
                        logging.info(f"Saved '{field_name}' with value '{value}' to column {date_column} for date {selected_date}.")

                # Save the workbook
                try:
                    save_workbook(self.workbook, self.file_path)
                    logging.info("Workbook saved successfully.")
                    messagebox.showinfo("Success", "Data saved successfully!")
                except Exception as e:
                    logging.error(f"Error saving workbook: {e}")
                    messagebox.showerror("Error", f"Failed to save workbook: {e}")

            elif active_tab == 1:  # Recognition Entry tab
                if hasattr(self, "recognition_manager"):
                    recognition_data = {k: v.get() for k, v in self.recognition_fields.items()}
                    self.recognition_manager.add_recognition(recognition_data, self.file_path)
                    logging.info(f"Recognition data saved: {recognition_data}")
                    messagebox.showinfo("Success", "Recognition data saved successfully!")
                else:
                    logging.error("Recognition manager not initialized.")
                    messagebox.showerror("Error", "Recognition manager not initialized.")

        except Exception as e:
            logging.error(f"Error saving data: {e}")
            messagebox.showerror("Error", f"Failed to save data: {e}")






if __name__ == "__main__":
    app = DataEntryApp()
    app.mainloop()
