from utils.excel_handler import get_sheet, write_to_cell, save_workbook
import logging
import os

# Configure logging for this module
# Log file in the app's directory
log_file_path = os.path.join(os.getcwd(), "error.log")
logging.basicConfig(
    filename=log_file_path,
    level=logging.ERROR,
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
)
logger = logging.getLogger(__name__)


class RecognitionEntryManager:
    def __init__(self, workbook):
        self.workbook = workbook
        # Ensure the correct sheet name is used
        self.sheet = get_sheet(workbook, "Recognitions")

    def add_recognition(self, recognition, file_path):
        """
        Add a new recognition entry to the sheet and save the workbook.
        :param recognition: Dictionary with keys `First Name`, `Last Name`, `Recognition`, and `Date`.
        :param file_path: The file path to save the workbook.
        """
        try:
            # Validate input
            required_fields = ["First Name",
                               "Last Name", "Recognition", "Date"]
            for field in required_fields:
                if field not in recognition or not recognition[field]:
                    raise ValueError(f"Missing required field: {field}")

            # Find the next empty row
            row = self.sheet.max_row + 1
            print(f"Adding data to row {row} in sheet {self.sheet.title}")

            # Write recognition data to the sheet
            write_to_cell(self.sheet, row, 1, recognition["First Name"])
            write_to_cell(self.sheet, row, 2, recognition["Last Name"])
            write_to_cell(self.sheet, row, 3, recognition["Recognition"])
            write_to_cell(self.sheet, row, 4, recognition["Date"])
            print(f"Successfully added recognition: {recognition}")

            # Save the workbook
            print(
                f"Saving workbook to {file_path} after adding recognition data...")
            save_workbook(self.workbook, file_path)
            print(f"Workbook successfully saved to {file_path}")
        except Exception as e:
            print(f"Error adding recognition: {e}")
            raise
