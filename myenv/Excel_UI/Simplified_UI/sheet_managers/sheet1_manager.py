from utils.excel_handler import get_sheet, write_to_cell
import logging

# Configure logging for this module
logger = logging.getLogger(__name__)

class Sheet1Manager:
    def __init__(self, workbook):
        try:
            self.workbook = workbook
            self.sheet = get_sheet(workbook, "Sheet 1")
        except Exception as e:
            logger.error(f"Error initializing Sheet1Manager: {e}")
            raise

    def save_metrics(self, metrics):
        """
        Save metrics data to the sheet.
        :param metrics: Dictionary of field names to values.
        """
        try:
            # Example mapping of fields to rows/columns in the sheet
            field_mapping = {
                "Days without Incident": (2, 2),
                "Haz ID's": (3, 2),
                "Safety Gemba Walk": (4, 2),
                "7S (Zone 26)": (5, 2),
                "7S (Zone 51)": (6, 2),
            }

            for field, value in metrics.items():
                if field not in field_mapping:
                    logger.warning(f"Unknown field: {field}. Skipping...")
                    continue

                row, col = field_mapping[field]
                write_to_cell(self.sheet, row, col, value)
            logger.info(f"Metrics saved successfully: {metrics}")
        except Exception as e:
            logger.error(f"Error saving metrics: {e}")
            raise
