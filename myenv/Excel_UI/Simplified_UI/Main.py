import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from sheet_managers.recognition_entry_manager import RecognitionEntryManager
from tkcalendar import DateEntry
from datetime import datetime, timedelta
from utils.excel_handler import (
    load_workbook, 
    save_workbook, 
    calculate_truck_fill_percentage, 
    save_days_without_incident_data, 
    load_days_without_incident_data, 
    extend_date_row, 
    find_or_add_date_column, 
    load_config,
    save_config, 
    CONFIG_FILE
    )

import json
import os
import logging



# Configure logging
logging.basicConfig(filename='app.log', level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

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

def save_window_size(self, event=None):
    """Save the current window size and position to config.json."""
    try:
        config = load_config()
        config["window_geometry"] = self.geometry()  # Save the current size and position
        save_config(config)
        logging.info(f"Window size and position saved: {config['window_geometry']}")
    except Exception as e:
        logging.error(f"Error saving window size: {e}")

def on_close(self):
    """Handle cleanup and save configuration on close."""
    self.save_window_size()
    self.destroy()  # Close the application


def auto_load_last_window_size(self):
    """Restore the last saved window size and position from config.json."""
    try:
        config = load_config()
        window_geometry = config.get("window_geometry")
        if window_geometry:
            self.geometry(window_geometry)
            logging.info(f"Restored window geometry: {window_geometry}")
    except Exception as e:
        logging.error(f"Error loading window size: {e}")


def update_days_without_incident(reset_toggle):
    try:
        data = load_days_without_incident_data()
        last_date = datetime.strptime(data["last_date"], "%Y-%m-%d").date()
        today = datetime.now().date()

        if reset_toggle:
            new_counter = 0
            logging.info("Days without Incident reset to 0.")
        else:
            days_passed = (today - last_date).days
            new_counter = data["counter"] + (days_passed if days_passed > 0 else 0)

        save_days_without_incident_data(new_counter, today)
        logging.info(f"Updated Days without Incident to {new_counter}.")
        return new_counter
    except Exception as e:
        logging.error(f"Error managing Days without Incident: {e}")
        return 0


class DataEntryApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Professional Data Entry")
        self.geometry("800x600")
        self.resizable(True, True)
        self.auto_load_last_window_size()
        self.bind("<Configure>", self.save_window_size)  # Save size on resize
        self.protocol("WM_DELETE_WINDOW", self.on_close)



        # Initialize workbook and other attributes
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

    def auto_load_last_window_size(self):
        """Restore the last saved window size and position from config.json."""
        try:
            config = load_config()  # Load the configuration from the JSON file
            window_geometry = config.get("window_geometry")
            if window_geometry:
                self.geometry(window_geometry)  # Apply the saved size and position
                logging.info(f"Restored window geometry: {window_geometry}")
        except Exception as e:
            logging.error(f"Error loading window size: {e}")

    def update_days_without_incident_live(self, *args):
        """Live update for Days without Incident field."""
        try:
            current_value = self.fields["Days without Incident"].get()
            if current_value.isdigit():
                new_counter = int(current_value)
                save_days_without_incident_data(new_counter, datetime.now())
                logging.info(f"Updated Days without Incident to {new_counter}.")
            else:
                logging.warning("Invalid input for Days without Incident.")
        except Exception as e:
            logging.error(f"Error in live update for Days without Incident: {e}")

    def update_days_without_incident_json(self, event):
        """Update JSON when Days without Incident loses focus."""
        try:
            current_value = self.fields["Days without Incident"].get()
            if current_value.isdigit():
                save_days_without_incident_data(int(current_value), datetime.now())
                logging.info("Days without Incident JSON updated.")
            else:
                logging.warning("Invalid input. Clearing Days without Incident field.")
                self.fields["Days without Incident"].set("")
        except Exception as e:
            logging.error(f"Error updating Days without Incident JSON: {e}")


    def create_ui(self):
        """Set up the main UI."""
        self.create_menu()
        self.create_scrollable_container()
        self.create_tabs()
        self.populate_fields()
        self.adjust_window_size()

    def create_menu(self):
        """Create the file menu."""
        file_button = ttk.Button(self, text="Open Excel File", command=self.load_excel_file)
        file_button.pack(pady=5)

    def create_scrollable_container(self):
        """Set up the scrollable container."""
        container = ttk.Frame(self)
        container.pack(fill="both", expand=True, padx=5, pady=5)

        self.canvas = tk.Canvas(container, highlightthickness=0)
        self.scrollbar = ttk.Scrollbar(container, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = ttk.Frame(self.canvas)

        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )
        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")

        # Enable mouse wheel scrolling
        def on_mouse_wheel(event):
            self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

        # Dynamically adjust the scroll region as the content resizes
        def update_scroll_region(event=None):
            self.canvas.configure(scrollregion=self.canvas.bbox("all"))
            # Check if the content fits without scrolling
            content_width = self.scrollable_frame.winfo_reqwidth()
            content_height = self.scrollable_frame.winfo_reqheight()
            window_width = self.winfo_width()
            window_height = self.winfo_height()

            # Deactivate scrollbar if content fits
            if content_width <= window_width and content_height <= window_height:
                self.scrollbar.pack_forget()  # Remove scrollbar
            else:
                self.scrollbar.pack(side="right", fill="y")  # Show scrollbar
                
            # Bind scrolling events
            self.canvas.bind_all("<MouseWheel>", on_mouse_wheel)  # For Windows and Linux
            self.canvas.bind_all("<Button-4>", lambda e: self.canvas.yview_scroll(-1, "units"))  # For macOS scroll up
            self.canvas.bind_all("<Button-5>", lambda e: self.canvas.yview_scroll(1, "units"))   # For macOS scroll down
        
    def create_tabs(self):
        """Create tabs for different sheets."""
        self.notebook = ttk.Notebook(self.scrollable_frame)
        self.notebook.pack(fill="both", expand=True, padx=5, pady=5)

        self.sheet1_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.sheet1_frame, text="Sheet 1")

        self.recognition_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.recognition_frame, text="Recognition Entry")

    def populate_fields(self):
        """Add fields to the Sheet 1 tab."""
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

        # Add date picker for selected date
        self.date_selection = tk.StringVar()
        date_picker = DateEntry(sheet1_group, textvariable=self.date_selection, width=20,
                                date_pattern="MM/dd/yyyy", background='darkblue', foreground='white', borderwidth=2)
        date_picker.grid(row=0, column=1, padx=10, pady=5, sticky="ew")
        ttk.Label(sheet1_group, text="Select Date:").grid(row=0, column=0, padx=10, pady=5, sticky="w")

        # Add reset days toggle
        self.reset_days_toggle = tk.BooleanVar(value=False)
        reset_button = ttk.Checkbutton(sheet1_group, text="Reset Days", variable=self.reset_days_toggle)
        reset_button.grid(row=1, column=2, padx=10, pady=5, sticky="ew")

        # Populate the Days without Incident field on UI load
        days_data = load_days_without_incident_data()  # Load data from config.json
        days_without_incident = days_data.get("counter", 0)
        self.fields["Days without Incident"].set(days_without_incident)
        self.fields["Days without Incident"].trace_add("write", self.update_days_without_incident_live)

        for idx, (label, var) in enumerate(self.fields.items()): 
            ttk.Label(sheet1_group, text=label).grid(row=idx, column=0, padx=10, pady=5, sticky="w")
            entry = ttk.Entry(sheet1_group, textvariable=var)
            entry.grid(row=idx, column=1, padx=10, pady=5, sticky="ew")    

            # Bind the Truck Fill % field to auto-calculate the percentage
            if label == "Truck Fill %":
                entry.bind("<FocusOut>", self.update_truck_fill_percentage)

        save_button = ttk.Button(self.scrollable_frame, text="Save Data", command=self.save_data)
        save_button.pack(pady=10)

        self.add_fields(self.recognition_frame, self.recognition_fields)

    def adjust_window_size(self):
        """Adjust the window size to fit all components while respecting screen size."""
        self.update_idletasks()  # Ensure all widgets are laid out and measured
        content_width = self.winfo_reqwidth()  # Width needed to fit all components
        content_height = self.winfo_reqheight()  # Height needed to fit all components

        # Get screen dimensions
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()

        # Ensure the window does not exceed the screen size
        new_width = min(content_width, screen_width)
        new_height = min(content_height, screen_height)

        # Set the window size to fit components or the screen size, whichever is smaller
        self.geometry(f"{new_width}x{new_height}")

    def save_window_size(self, event=None):
        """Save the current window size and position to config.json."""
        try:
            config = load_config()
            config["window_geometry"] = self.geometry()  # Save the current size and position
            save_config(config)
            logging.info(f"Window size and position saved: {config['window_geometry']}")
        except Exception as e:
            logging.error(f"Error saving window size: {e}")

    def on_close(self):
        """Handle cleanup and save configuration on close."""
        self.save_window_size()
        self.destroy()  # Close the application


    def update_truck_fill_percentage(self, event):
        """Automatically calculate and update the Truck Fill % field."""
        try:
            truck_fill_value = self.fields["Truck Fill %"].get()
            if truck_fill_value:
                percentage = calculate_truck_fill_percentage(truck_fill_value)
                self.fields["Truck Fill %"].set(percentage)  # Update the field with the calculated percentage
                logging.info(f"Updated Truck Fill % to {percentage}")
        except ValueError as e:
            logging.error(f"Error updating Truck Fill %: {e}")
            messagebox.showerror("Error", "Invalid input for Truck Fill %. Please enter a numeric value between 0 and 26.")

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
            # Get the selected date
            selected_date = self.date_selection.get()
            if not selected_date:
                messagebox.showerror("Error", "Please select a date.")
                return

            selected_date = datetime.strptime(selected_date, "%m/%d/%Y").date()
            data_sheet = self.sheet_mapping.get("Data")

            if data_sheet:
                # Find or add the column for the selected date
                date_column = find_or_add_date_column(data_sheet, selected_date, start_column=3)

                # Save field data to the Excel sheet
                for field_name, (sheet_name, (row, _)) in self.field_to_sheet_mapping.items():
                    if sheet_name == "Data":
                        value = self.fields[field_name].get()

                        # Handle Truck Fill % separately
                        if field_name == "Truck Fill %":
                            try:
                                # Strip '%' if present and convert to numeric
                                if "%" in value:
                                    value = float(value.strip('%')) / 100 * 26
                                # Recalculate percentage for consistency
                                value = calculate_truck_fill_percentage(value)
                            except ValueError as e:
                                logging.error(f"Error saving Truck Fill %: {e}")
                                messagebox.showerror("Error", "Invalid Truck Fill % value. Please correct it.")
                                return

                        # Write the value to the Excel sheet
                        data_sheet.cell(row=row, column=date_column).value = value
                        logging.info(f"Saved '{field_name}' with value '{value}' to column {date_column}.")

                # Save the workbook
                save_workbook(self.workbook, self.file_path)
                logging.info("Workbook saved successfully.")
                messagebox.showinfo("Success", "Data saved successfully!")

        except Exception as e:
            logging.error(f"Error saving data: {e}")
            messagebox.showerror("Error", f"Failed to save data: {e}")



if __name__ == "__main__":
    app = DataEntryApp()
    app.mainloop()
