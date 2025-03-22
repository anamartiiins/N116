import sys
import xlwings as xw
from src.constants import PATH_EXCEL
from src.extract.extract import get_excel_metadata

def main():
    """Main function to process Excel operations based on VBA input."""
    if len(sys.argv) < 2:
        print("No function specified.")
        return

    operation = sys.argv[1]  # First argument is the function name

    # Open Excel workbook
    wb = xw.Book(PATH_EXCEL)
    sheet = wb.sheets.active

    # Handle different operations
    if operation == "get_metadata":
        metadata = get_excel_metadata(sheet)
        print(metadata)

    else:
        print(f"Unknown operation: {operation}")

    # Save and close
    wb.save()
    wb.close()

if __name__ == "__main__":
    main()
