from utils.excel_handler import get_sheet, write_to_cell
import logging

# Configure logging for this module
logger = logging.getLogger(__name__)

class RecognitionEntryManager:
    def __init__(self, workbook):
        try:
            self.workbook = workbook
            self.sheet = get_sheet(workbook, "Recognition Entry")
        except Exception as e:
            logger.error(f"Error initializing RecognitionEntryManager: {e}")
            raise

    def add_recognition(self, recognition):
        """
        Add a new recognition entry to the sheet.
        :param recognition: Dictionary with keys `First Name`, `Last Name`, `Recognition`, and `Date`.
        """
        try:
            # Validate input
            required_fields = ["First Name", "Last Name", "Recognition", "Date"]
            for field in required_fields:
                if field not in recognition or not recognition[field]:
                    logger.error(f"Missing field: {field} in recognition entry.")
                    raise ValueError(f"Missing required field: {field}")

            # Find the next empty row
            row = self.sheet.max_row + 1

            # Write recognition data to the sheet
            write_to_cell(self.sheet, row, 1, recognition["First Name"])
            write_to_cell(self.sheet, row, 2, recognition["Last Name"])
            write_to_cell(self.sheet, row, 3, recognition["Recognition"])
            write_to_cell(self.sheet, row, 4, recognition["Date"])
            logger.info(f"Recognition added successfully: {recognition}")
        except Exception as e:
            logger.error(f"Error adding recognition: {e}")
            raise
