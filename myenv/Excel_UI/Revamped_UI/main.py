import tkinter as tk
from tkinter import filedialog

def open_file():
    file_path = filedialog.askopenfilename(title="Select Excel File", filetypes=[("Excel files", "*.xlsx")])
    if file_path:
        print(f"Selected file: {file_path}")

def submit_data():
    print("Submit button clicked")

# Initialize the main window
root = tk.Tk()
root.title("Data Entry UI")
root.geometry("400x400")  # Set the window size

# Create a frame to hold all widgets
main_frame = tk.Frame(root)
main_frame.pack(fill="both", expand=True)

# Create a Canvas and a Scrollbar
canvas = tk.Canvas(main_frame)
scrollbar = tk.Scrollbar(main_frame, orient="vertical", command=canvas.yview)
scrollable_frame = tk.Frame(canvas)

# Configure the scrollable frame
scrollable_frame.bind(
    "<Configure>",
    lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
)

canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
canvas.configure(yscrollcommand=scrollbar.set)

# Add the scrollbar to the right side of the main frame
canvas.pack(side="left", fill="both", expand=True)
scrollbar.pack(side="right", fill="y")

# Button to select Excel file
file_button = tk.Button(scrollable_frame, text="Select Excel File", command=open_file)
file_button.grid(row=0, column=0, columnspan=2, pady=(0, 20))

# Labels and Entry fields for data input
labels = [
    "Days without Incident", "Haz ID's", "Safety Gemba Walk", "7S (Zone 26)", "7S (Zone 51)", "Errors", "PCD Returns",
    "Jobs on Hold", "Productivity", "OTIF %", "Huddles", "Truck Fill %", "Recognitions", "MC Compliance",
    "Cost Savings", "Rever's", "Project's"
]

entry_fields = []

for i, label in enumerate(labels):
    tk.Label(scrollable_frame, text=label).grid(row=i+1, column=0, sticky="e", padx=5, pady=5)
    entry = tk.Entry(scrollable_frame)
    entry.grid(row=i+1, column=1, padx=5, pady=5)
    entry_fields.append(entry)

# Submit button at the bottom
submit_button = tk.Button(scrollable_frame, text="Submit", command=submit_data)
submit_button.grid(row=len(labels) + 1, column=0, columnspan=2, pady=20)

# Enable scrolling with the mouse wheel
def on_mouse_wheel(event):
    canvas.yview_scroll(int(-1*(event.delta/120)), "units")

canvas.bind_all("<MouseWheel>", on_mouse_wheel)

# Run the application
root.mainloop()
