import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
from utils.excel_handler import load_workbook, save_workbook, calculate_truck_fill_percentage
from sheet_managers.recognition_entry_manager import RecognitionEntryManager
from tkcalendar import DateEntry
from datetime import datetime, timedelta
import json
import os

CONFIG_FILE = "config.json"  # File to store the last file path

print(calculate_truck_fill_percentage(13))  # Should print 50.0

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
        self.date_selection = tk.StringVar()


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

    def adjust_window_size(self):
        """Adjust the window size based on the content."""
        self.update_idletasks()  # Ensure all geometry calculations are updated
        width = self.notebook.winfo_reqwidth() + 20  # Add padding
        height = self.notebook.winfo_reqheight() + 20
        self.geometry(f"{width}x{height}")

    def create_ui(self):
        """Create the main UI."""
        # File selection button
        file_button = ttk.Button(self, text="Open Excel File", command=self.load_excel_file)
        file_button.pack(pady=5)  # Reduce padding to tighten layout

        # Create a container for the scrollable area
        container = ttk.Frame(self)
        container.pack(fill="both", expand=True, padx=5, pady=5)  # Adjust padding here

        # Create canvas and scrollbar
        canvas = tk.Canvas(container)
        scrollbar = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)

        # Configure the canvas and scrollbar
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        # Pack the canvas and scrollbar
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # Enable mouse wheel scrolling
        def on_mouse_wheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

        canvas.bind_all("<MouseWheel>", on_mouse_wheel)  # Windows/Linux
        canvas.bind_all("<Button-4>", lambda e: canvas.yview_scroll(-1, "units"))  # macOS scroll up
        canvas.bind_all("<Button-5>", lambda e: canvas.yview_scroll(1, "units"))   # macOS scroll down

        # Add a Notebook for tabs inside the scrollable frame
        self.notebook = ttk.Notebook(scrollable_frame)
        self.notebook.pack(fill="both", expand=True, padx=5, pady=5)  # Adjust padding here

        # Tabs for Sheet 1 and Recognition Entry
        self.sheet1_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.sheet1_frame, text="Sheet 1")

        self.recognition_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.recognition_frame, text="Recognition Entry")

        # Define field mappings
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
        sheet1_group.pack(fill="both", expand=True, padx=5, pady=5)  # Adjust padding here

        # Add a DateEntry for selecting the date
        self.date_selection = tk.StringVar()
        date_picker = DateEntry(sheet1_group, textvariable=self.date_selection, width=20,
                                date_pattern="MM/dd/yyyy", background='darkblue', foreground='white', borderwidth=2)
        date_picker.grid(row=0, column=1, padx=10, pady=5, sticky="ew")
        ttk.Label(sheet1_group, text="Select Date:").grid(row=0, column=0, padx=10, pady=5, sticky="w")

        # Toggle for Days without Incident
        self.reset_days_toggle = tk.BooleanVar(value=False)
        reset_button = ttk.Checkbutton(sheet1_group, text="Reset Days", variable=self.reset_days_toggle)
        reset_button.grid(row=1, column=2, padx=10, pady=5, sticky="ew")

        # Add fields to the LabelFrame
        for idx, (label, var) in enumerate(self.fields.items(), start=1):  # Start after the date picker
            ttk.Label(sheet1_group, text=label).grid(row=idx, column=0, padx=10, pady=5, sticky="w")
            ttk.Entry(sheet1_group, textvariable=var).grid(row=idx, column=1, padx=10, pady=5, sticky="ew")
            sheet1_group.columnconfigure(1, weight=1)

        # Add fields to the Recognition Entry tab
        self.add_fields(self.recognition_frame, self.recognition_fields)

        # Add Save Data button at the bottom
        save_button = ttk.Button(scrollable_frame, text="Save Data", command=self.save_data)
        save_button.pack(pady=10)

        # Bind the Enter key to the Save Data button
        save_button.bind("<Return>", lambda event: self.save_data())

        # Adjust window size dynamically after UI is created
        self.adjust_window_size()

    def add_fields(self, frame, fields):
        """Add labeled input fields for a given set of fields."""
        for idx, (field_name, var) in enumerate(fields.items()):
            ttk.Label(frame, text=field_name).grid(row=idx, column=0, padx=10, pady=5, sticky="w")
            ttk.Entry(frame, textvariable=var).grid(row=idx, column=1, padx=10, pady=5, sticky="ew")
            frame.columnconfigure(1, weight=1)

    def manage_days_without_incident(self):
        """Manage the Days without Incident counter."""
        # Load the current counter and last update date
        config_file = "days_without_incident.json"
        if not os.path.exists(config_file):
            data = {"counter": 0, "last_date": datetime.now().strftime("%Y-%m-%d")}
        else:
            with open(config_file, "r") as f:
                data = json.load(f)

        # Parse the last date and calculate days passed
        last_date = datetime.strptime(data["last_date"], "%Y-%m-%d")
        today = datetime.now().date()

        # Check if toggle is on
        if self.reset_days_toggle.get():
            # Reset the counter to 0 if toggle is selected
            data["counter"] = 0
            print("Resetting Days without Incident to 0.")
        else:
            # Use a custom value if provided
            custom_value = self.fields["Days without Incident"].get()
            if custom_value.isdigit():
                data["counter"] = int(custom_value)
                print(f"Setting Days without Incident to custom value: {custom_value}.")
            else:
                # Increment the counter based on days passed
                days_passed = (today - last_date).days
                if days_passed > 0:
                    data["counter"] += days_passed
                    print(f"Incrementing Days without Incident by {days_passed} days.")

        # Save the updated counter and last date
        data["last_date"] = today.strftime("%Y-%m-%d")
        with open(config_file, "w") as f:
            json.dump(data, f)

        print(f"Days without Incident: {data['counter']} (Last updated: {data['last_date']})")
        return data["counter"]

    def load_excel_file(self):
        """Load the Excel file and remember its path."""
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            try:
                self.workbook = load_workbook(file_path)
                self.file_path = file_path  # Store the selected file path
                save_last_file_path(file_path)  # Save the file path for future use
                self.sheet_mapping = {name: self.workbook[name] for name in self.workbook.sheetnames}

                if "Recognitions" in self.sheet_mapping:
                    print("Initializing RecognitionEntryManager...")
                    self.recognition_manager = RecognitionEntryManager(self.workbook)
                    print("RecognitionEntryManager initialized successfully.")

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
                selected_date = self.date_selection.get()
                if not selected_date:
                    tk.messagebox.showerror("Error", "Please select a date.")
                    return

                # Handle Days without Incident logic
                days_without_incident = self.manage_days_without_incident()
                self.fields["Days without Incident"].set(days_without_incident)

                data_sheet = self.sheet_mapping.get("Data")
                if not data_sheet:
                    tk.messagebox.showerror("Error", "Data sheet not found in the workbook.")
                    return

                # Find the column corresponding to the selected date
                selected_date = self.date_selection.get()
                date_column = None
                for col in range(2, data_sheet.max_column + 1):
                    cell_value = data_sheet.cell(row=1, column=col).value
                    # Convert cell_value to string in the same format
                    if cell_value and isinstance(cell_value, datetime):
                        cell_value = cell_value.strftime("%m/%d/%Y")  # Convert datetime to MM/DD/YYYY
                    if cell_value == selected_date:
                        date_column = col
                        break

                if not date_column:
                    tk.messagebox.showerror("Error", f"Selected date '{selected_date}' not found in the Data sheet.")
                    return

                # Handle "Truck Fill %" calculation
                truck_fill_field = self.fields.get("Truck Fill %")
                if truck_fill_field:
                    try:
                        entered_value = truck_fill_field.get()
                        percentage = calculate_truck_fill_percentage(entered_value)  # Now includes the '%' symbol
                        truck_fill_field.set(percentage)  # Update the field with the formatted percentage
                        print(f"Calculated Truck Fill %: {percentage}")
                    except ValueError as e:
                        tk.messagebox.showerror("Error", str(e))
                        return

                # Save data to the correct column under the selected date
                for field_name, (sheet_name, (row, _)) in self.field_to_sheet_mapping.items():
                    value = self.fields[field_name].get()  # Get user input
                    if sheet_name == "Data":
                        data_sheet.cell(row=row, column=date_column).value = value
                        print(f"Field '{field_name}' -> Sheet '{sheet_name}', Cell ({row},{date_column}): '{value}'")
            
                if self.file_path:
                    try:
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

            elif active_tab == 1:  # Recognition Entry tab
                if hasattr(self, "recognition_manager") and self.recognition_manager:
                    recognition_data = {k: v.get() for k, v in self.recognition_fields.items()}
                    self.recognition_manager.add_recognition(recognition_data, self.file_path)
                    print(f"Recognition data saved: {recognition_data}")
                    tk.messagebox.showinfo("Success", "Recognition data saved successfully!")
                else:
                    tk.messagebox.showerror("Error", "Recognition manager not initialized. Please load a valid Excel file.")

        except Exception as e:
            print(f"Error saving data: {e}")
            tk.messagebox.showerror("Error", f"Failed to save data: {e}")


if __name__ == "__main__":
    app = DataEntryApp()
    app.mainloop()
