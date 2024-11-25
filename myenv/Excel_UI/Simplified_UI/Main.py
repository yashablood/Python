import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
from utils.excel_handler import load_workbook, save_workbook
from sheet_managers.sheet1_manager import Sheet1Manager
from sheet_managers.recognition_entry_manager import RecognitionEntryManager


class DataEntryApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Professional Data Entry")
        self.geometry("800x600")
        self.resizable(True, True)

        self.workbook = None
        self.sheet1_manager = None
        self.recognition_manager = None

        # Initialize UI
        self.create_ui()

    def create_ui(self):
        # Add file selection button
        file_button = ttk.Button(self, text="Open Excel File", command=self.load_excel_file)
        file_button.pack(pady=10)

        # Notebook
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill="both", expand=True, padx=10, pady=10)

        # Tabs
        self.sheet1_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.sheet1_frame, text="Sheet 1")

        self.recognition_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.recognition_frame, text="Recognition Entry")

        # Add Fields for Sheet 1
        self.sheet1_fields = {
            "Days without Incident": tk.StringVar(),
            "Haz ID's": tk.StringVar(),
            "Safety Gemba Walk": tk.StringVar(),
            "7S (Zone 26)": tk.StringVar(),
            "7S (Zone 51)": tk.StringVar(),
        }
        self.add_fields(self.sheet1_frame, self.sheet1_fields, self.save_sheet1_data)

        # Add Recognitions Section for Sheet 2
        self.recognition_fields = {
            "First Name": tk.StringVar(),
            "Last Name": tk.StringVar(),
            "Recognition": tk.StringVar(),
            "Date": tk.StringVar(),
        }
        self.add_fields(self.recognition_frame, self.recognition_fields, self.add_recognition_data)

    def add_fields(self, frame, fields, submit_callback):
        """Add labeled entry fields and a submit button to the frame."""
        for idx, (label, var) in enumerate(fields.items()):
            ttk.Label(frame, text=label).grid(row=idx, column=0, padx=10, pady=5, sticky="w")
            ttk.Entry(frame, textvariable=var).grid(row=idx, column=1, padx=10, pady=5, sticky="ew")
            frame.columnconfigure(1, weight=1)

        ttk.Button(frame, text="Submit", command=submit_callback).grid(
            row=len(fields), column=0, columnspan=2, pady=10
        )

    def load_excel_file(self):
        """Load the Excel file and initialize managers."""
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            self.workbook = load_workbook(file_path)
            self.sheet1_manager = Sheet1Manager(self.workbook)
            self.recognition_manager = RecognitionEntryManager(self.workbook)
            print("Excel file loaded successfully!")

    def save_sheet1_data(self):
        """Save data for Sheet 1."""
        if self.sheet1_manager:
            self.sheet1_manager.save_metrics({k: v.get() for k, v in self.sheet1_fields.items()})
            print("Sheet 1 data saved!")

    def add_recognition_data(self):
        """Add a recognition entry."""
        if self.recognition_manager:
            self.recognition_manager.add_recognition({k: v.get() for k, v in self.recognition_fields.items()})
            print("Recognition entry added!")

    def save_and_close(self):
        """Save changes and close the app."""
        if self.workbook:
            save_workbook(self.workbook, "output.xlsx")
            print("Workbook saved!")
        self.quit()


if __name__ == "__main__":
    app = DataEntryApp()
    app.mainloop()
