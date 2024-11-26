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

    # Attempt to load the last used file
    last_file = load_last_file_path()
    if last_file:
        try:
            self.workbook = load_workbook(last_file)
            self.file_path = last_file
            self.sheet_mapping = {name: self.workbook[name] for name in self.workbook.sheetnames}
            print(f"Auto-loaded file: {last_file}")
        except Exception as e:
            print(f"Failed to auto-load file: {e}")

    # Initialize UI
    self.create_ui()
