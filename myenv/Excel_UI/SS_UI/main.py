import tkinter as tk
from tkinter import filedialog

class StandardStartUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Standard Start Entry")
        self.file_path = None

        # Excel file selection
        self.file_button = tk.Button(root, text="Select Excel File", command=self.select_file)
        self.file_button.pack(pady=10)

        # Input frame
        self.input_frame = tk.Frame(root)
        self.input_frame.pack(pady=10)

        self.labels = ["Name", "Standard Start Completed"]
        self.entries = {}

        for label in self.labels:
            row = tk.Frame(self.input_frame)
            row.pack(fill="x", pady=5)

            tk.Label(row, text=label, width=25, anchor="w").pack(side="left")
            entry = tk.Entry(row, width=40)
            entry.pack(side="right", expand=True, fill="x")
            self.entries[label] = entry

        # Submit button
        self.submit_button = tk.Button(root, text="Submit", command=self.submit_data)
        self.submit_button.pack(pady=20)

    def select_file(self):
        self.file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
        if self.file_path:
            print(f"Selected file: {self.file_path}")

    def submit_data(self):
        # Placeholder for future Excel logic
        data = {label: entry.get() for label, entry in self.entries.items()}
        print("Submitted Data:", data)

if __name__ == "__main__":
    root = tk.Tk()
    app = StandardStartUI(root)
    root.mainloop()
