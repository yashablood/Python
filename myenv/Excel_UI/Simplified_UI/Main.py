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
DAYS_WITHOUT_INCIDENT_FILE = os.path.abspath("days_without_incident.json")
#DAYS_WITHOUT_INCIDENT_FILE = "days_without_incident.json"


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

        # Add all missing dates until the current date
        while next_date <= today:
            # Check if the current date already exists in the date row
            is_duplicate = any(
                sheet.cell(row=1, column=col).value == next_date
                for col in range(start_column, sheet.max_column + 1)
            )
            if not is_duplicate:
                cell = sheet.cell(row=1, column=current_column)
                cell.value = next_date
                cell.number_format = "dd-mmm"  # Ensure consistent formatting
                logging.info(f"Added date {next_date.strftime('%d-%b')} to column {current_column}.")
                current_column += 1
            next_date += timedelta(days=1)

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

    def update_days_without_incident_json(self, event):
        """Update the JSON file with the manually entered 'Days without Incident' value."""
        try:
            current_value = self.fields["Days without Incident"].get()
            if current_value.isdigit():
                new_counter = int(current_value)
                data = {"counter": new_counter, "last_date": datetime.now().strftime("%Y-%m-%d")}
                
                # Create the file if it doesn't exist
                if not os.path.exists(DAYS_WITHOUT_INCIDENT_FILE):
                    logging.info(f"JSON file not found. Creating new file at: {DAYS_WITHOUT_INCIDENT_FILE}")
                    with open(DAYS_WITHOUT_INCIDENT_FILE, "w") as f:
                        json.dump(data, f)
                else:
                    with open(DAYS_WITHOUT_INCIDENT_FILE, "w") as f:
                        json.dump(data, f)

                logging.info(f"Days without Incident JSON updated: Counter = {new_counter}, Last Date = {datetime.now().strftime('%Y-%m-%d')}")
            else:
                messagebox.showerror("Error", "Invalid input for 'Days without Incident'. Please enter a number.")
        except Exception as e:
            logging.error(f"Failed to update Days without Incident JSON: {e}")
            messagebox.showerror("Error", "Failed to update Days without Incident JSON.")



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

        # Bind live update for "Days without Incident" field
        days_field = ttk.Entry(sheet1_group, textvariable=self.fields["Days without Incident"])
        days_field.grid(row=1, column=1, padx=10, pady=5, sticky="ew")
        days_field.bind("<FocusOut>", lambda e: print("FocusOut triggered"),  self.update_days_without_incident_json)

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

                logging.info(f"Loaded file: {file_path}")
            except Exception as e:
                logging.error(f"Error loading Excel file: {e}")
                messagebox.showerror("Error", f"Failed to load Excel file: {e}")


    def save_data(self):
        """Save data to the workbook."""
        if not self.workbook:
            messagebox.showerror("Error", "No workbook loaded.")
            return

        try:
            # Get the manually entered value for "Days without Incident"
            manual_days_without_incident = self.fields["Days without Incident"].get()
            if manual_days_without_incident.isdigit():  # Ensure it's a valid number
                manual_counter = int(manual_days_without_incident)

                # Update the JSON file with the manually entered value
                try:
                    data = {"counter": manual_counter, "last_date": datetime.now().strftime("%Y-%m-%d")}
                    with open(DAYS_WITHOUT_INCIDENT_FILE, "w") as f:
                        json.dump(data, f)
                    logging.info(f"Updating JSON file at: {DAYS_WITHOUT_INCIDENT_FILE}")
                    logging.info(f"Updated JSON with manually entered Days without Incident: {manual_counter}")


                except Exception as e:
                    logging.error(f"Failed to update Days without Incident JSON: {e}")
                    messagebox.showerror("Error", "Failed to update Days without Incident JSON.")

            active_tab = self.notebook.index(self.notebook.select())

            # Ensure the date row is complete
            data_sheet = self.sheet_mapping.get("Data")
            if data_sheet:
                logging.info("Ensuring the date row is complete.")
                
                # Call extend_date_row to add missing dates
                extend_date_row(data_sheet, start_column=3)

                # Save the workbook to ensure changes are applied
                save_workbook(self.workbook, self.file_path)
                logging.info("Workbook saved after extending the date row.")

                # Reload the workbook to refresh the sheet mapping
                self.workbook = load_workbook(self.file_path)
                self.sheet_mapping = {name: self.workbook[name] for name in self.workbook.sheetnames}
                data_sheet = self.sheet_mapping.get("Data")

                # Get the selected date from the UI
                selected_date = self.date_selection.get()
                if not selected_date:
                    messagebox.showerror("Error", "Please select a date.")
                    return

                try:
                    # Convert the selected date to a datetime object
                    selected_date = datetime.strptime(selected_date, "%m/%d/%Y").date()
                except ValueError as e:
                    logging.error(f"Invalid date selected: {e}")
                    messagebox.showerror("Error", f"Invalid date format: {selected_date}")
                    return

                # Find the column corresponding to the selected date
                date_column = None
                for col in range(3, data_sheet.max_column + 1):  # Start from column 3
                    cell_value = data_sheet.cell(row=1, column=col).value
                    if cell_value and isinstance(cell_value, datetime) and cell_value.date() == selected_date:
                        date_column = col
                        break

                # If the date is still not found (which shouldn't happen), raise an error
                if not date_column:
                    logging.error(f"Date {selected_date.strftime('%d-%b')} not found after extending date row.")
                    messagebox.showerror("Error", f"Date {selected_date.strftime('%d-%b')} not found in the date row.")
                    return

                # Handle Truck Fill % calculation
                truck_fill_field = self.fields.get("Truck Fill %")
                if truck_fill_field:
                    try:
                        entered_value = truck_fill_field.get()  # Get user input
                        # Calculate and validate the percentage
                        percentage = calculate_truck_fill_percentage(entered_value)
                        truck_fill_field.set(percentage)  # Update the field with the formatted percentage
                        logging.info(f"Calculated Truck Fill %: {percentage}")
                    except ValueError as e:
                        logging.error(f"Truck Fill % calculation error: {e}")
                        messagebox.showerror("Error", str(e))
                        return

                # Save data to the appropriate column for the selected date
                for field_name, (sheet_name, (row, _)) in self.field_to_sheet_mapping.items():
                    if sheet_name == "Data":
                        value = self.fields[field_name].get()  # Get the value from the UI field
                        data_sheet.cell(row=row, column=date_column, value=value)
                        logging.info(f"Updated '{field_name}' with value '{value}' in column {date_column} for date {selected_date}.")

            elif active_tab == 1:  # Recognition Entry tab
                if hasattr(self, "recognition_manager"):
                    recognition_data = {k: v.get() for k, v in self.recognition_fields.items()}
                    self.recognition_manager.add_recognition(recognition_data, self.file_path)
                    logging.info(f"Recognition data saved: {recognition_data}")
                    messagebox.showinfo("Success", "Recognition data saved successfully!")
                else:
                    logging.error("Recognition manager not initialized.")
                    messagebox.showerror("Error", "Recognition manager not initialized.")

            # Save the workbook after all changes
            save_workbook(self.workbook, self.file_path)
            logging.info("Workbook saved successfully.")
            messagebox.showinfo("Success", "Data saved successfully!")

        except Exception as e:
            logging.error(f"Error saving data: {e}")
            messagebox.showerror("Error", f"Failed to save data: {e}")


if __name__ == "__main__":
    app = DataEntryApp()
    app.mainloop()
