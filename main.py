# General imports
import argparse
import sys
import xlwings as xw

# Internal imports
from src.constants import PATH_EXCEL, FORMULAS_MAPPING_ARTICLES
from src.extract import get_excel_metadata
from src.process import dynamic_formulas_mapping, col_idx_to_letter, insert_product_rows, insert_zone_row, \
    add_or_delete_row_between_columns, create_supplier_sheets


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

    parser.add_argument(
        "extras", nargs="*", help="Extra arguments: operation, row_numbers, etc."
    )
    return parser.parse_args()


def main(mode: str, extras: list):
    """
    Args:
        mode (str): Mode to run the script in: development or production
    """
    if mode == "production":
        # operation = sys.argv[2]
        # row_numbers = [int(r.strip()) for r in sys.argv[3].split(",")]
        # zone_name_arg = sys.argv[4] if len(sys.argv) > 4 else None
        operation = extras[0] if len(extras) > 0 else None
        row_numbers = [int(r.strip()) for r in extras[1].split(",")] if len(extras) > 1 else []
        name_supplier = extras[2] if len(extras) > 2 else None
        zone_name_arg = extras[3] if len(extras) > 3 else None
        print(sys.argv)

    elif mode == "development":
        operation = "insert_zone_row"
        row_numbers = [15]
        name_supplier = "MINDOL"
        zone_name_arg = "Quarto"

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

    if operation == "insert_product_rows":
        insert_product_rows(
            sheet=sheet,
            row_numbers=row_numbers,
            column_indices=column_indices,
            formula_mapping=formulas_mapping_articles,
        )

    elif operation == "insert_zone_row":
        insert_zone_row(sheet=sheet,
                        row_numbers=row_numbers,
                        column_indices=column_indices,
                        zone_name=zone_name_arg)

    elif operation == "create_supplier_sheets":
        create_supplier_sheets(sheet=sheet,
                               only_supplier=name_supplier)
    else:
        print(f"Unknown operation: {operation}")


if __name__ == "__main__":
    args = parse_args()
    mode = args.mode
    extras = args.extras
    main(mode=mode, extras=extras)
