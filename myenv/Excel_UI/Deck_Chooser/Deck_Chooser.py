import tkinter as tk
from tkinter import messagebox
import json
import random
import os

# File to save entries
script_dir = os.path.dirname(os.path.abspath(__file__))
DATA_FILE = os.path.join(script_dir, "entries.json")

# Load existing data
def load_data():
    if os.path.exists(DATA_FILE):
        with open(DATA_FILE, "r") as file:
            return json.load(file)
    return []

# Save data to JSON
def save_data(data):
    with open(DATA_FILE, "w") as file:
        json.dump(data, file, indent=4)

# Add new entry
def add_entry():
    entry = entry_input.get().strip()
    if entry:
        entries.append(entry)
        save_data(entries)
        listbox.insert(tk.END, entry)
        entry_input.delete(0, tk.END)
    else:
        messagebox.showwarning("Input Error", "Please enter something.")

# Select random entry
def select_random():
    if entries:
        random_entry = random.choice(entries)
        messagebox.showinfo("Random Selection", f"Random Entry: {random_entry}")
    else:
        messagebox.showwarning("No Entries", "No entries available to select.")

# Initialize main window
root = tk.Tk()
root.title("JSON Entry Manager")

# Entry Input
tk.Label(root, text="Enter Something:").grid(row=0, column=0, padx=10, pady=5, sticky="w")
entry_input = tk.Entry(root, width=30)
entry_input.grid(row=0, column=1, padx=10, pady=5)

# Buttons
add_button = tk.Button(root, text="Add Entry", command=add_entry)
add_button.grid(row=0, column=2, padx=10, pady=5)

random_button = tk.Button(root, text="Select Random", command=select_random)
random_button.grid(row=1, column=1, padx=10, pady=5)

# Listbox to show entries
tk.Label(root, text="Entries:").grid(row=2, column=0, padx=10, pady=5, sticky="nw")
listbox = tk.Listbox(root, width=50, height=10)
listbox.grid(row=2, column=1, columnspan=2, padx=10, pady=5, sticky="w")

# Load existing entries into the listbox
entries = load_data()
for entry in entries:
    listbox.insert(tk.END, entry)

# Start the application
root.mainloop()
