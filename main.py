# General imports
import argparse
import sys
import xlwings as xw

# Internal imports
from src.constants import PATH_EXCEL, FORMULAS_MAPPING_ARTICLES, INPUT_CELL_TRANSPORT_TOTAL
from src.extract import get_excel_metadata
from src.process import dynamic_formulas_mapping, col_idx_to_letter, insert_product_between_columns, \
    add_or_delete_row_between_columns, find_cell_by_content

def parse_args():
    parser = argparse.ArgumentParser(
        description="Excel Processor",
        usage="%(prog)s [-m development|production]",
        allow_abbrev=False,
    )
    parser.add_argument(
        "-m", "--mode",
        type=str,
        default="development",
        choices=["development", "production"],
        help="Choose the mode to run: development or production",
    )
    return parser.parse_args()


def main(mode: str):
    """
    Args:
        mode (str): Mode to run the script in: development or production
    """
    if mode == "production":
        operation = sys.argv[2]
        row_numbers = [int(r.strip()) for r in sys.argv[3].split(",")]
        zone_name_arg = sys.argv[4] if len(sys.argv) > 4 else None

    elif mode == "development":
        operation = "insert_product_between_columns"
        row_numbers = [12, 13]
        zone_name_arg = None

    # Create connection to excel
    wb = xw.Book(PATH_EXCEL)
    sheet = wb.sheets.active

    # Extract general and important data
    metadata = get_excel_metadata(sheet)

    column_indices = {
        header: {"index": idx + 2, "column": col_idx_to_letter(idx + 2)}
        for idx, header in enumerate(metadata['headers'])
    }

    formulas_mapping_articles = dynamic_formulas_mapping(FORMULAS_MAPPING_ARTICLES, column_indices)

    # find_cell_by_content(sheet=sheet, content=INPUT_CELL_TRANSPORT_TOTAL, return_type=)

    if operation == "insert_product_between_columns":
        insert_product_between_columns(
            sheet=sheet,
            row_numbers=row_numbers,
            column_indices=column_indices,
            formula_mapping=formulas_mapping_articles,
            start_column=metadata['headers'][0],
            end_column=metadata['headers'][-1],
        )

    elif operation == "delete_between_columns":
        add_or_delete_row_between_columns(
            sheet=sheet,
            row_numbers=row_numbers,
            column_indices=column_indices,
            start_column=metadata['headers'][0],
            end_column=metadata['headers'][-1],
            action="delete",
        )

    elif operation == "add_zone":
        add_or_delete_row_between_columns(
            sheet=sheet,
            row_numbers=row_numbers,
            column_indices=column_indices,
            start_column=metadata['headers'][0],
            end_column=metadata['headers'][-1],
            action="add",
            zone="yes",
            zone_name=zone_name_arg,
        )

    else:
        print(f"Unknown operation: {operation}")


if __name__ == "__main__":
    args = parse_args()
    mode = args.mode
    main(mode=mode)

