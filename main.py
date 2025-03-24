import sys
import xlwings as xw
from src.constants import PATH_EXCEL, FORMULAS_MAPPING_ARTICLES
from src.extract import get_excel_metadata
from src.process import dynamic_formulas_mapping, col_idx_to_letter, insert_product_between_columns, add_or_delete_row_between_columns

# def main():
#     """Main function to process Excel operations based on VBA input."""
#     if len(sys.argv) < 2:
#         print("No function specified.")
#         return
#
#     operation = sys.argv[1]
#
#     # Connect to the workbook and get the active sheet
#     wb = xw.Book(path_excel)
#     sheet = wb.sheets.active
#
#     # Handle different operations
#     if operation == "get_metadata":
#         metadata = get_excel_metadata(sheet)
#         print(metadata)
#
#     else:
#         print(f"Unknown operation: {operation}")
#
#     # Save and close
#     wb.save()
#     wb.close()
if __name__ == "__main__":
    import sys

    operation = sys.argv[1]
    row_numbers_arg = sys.argv[2]
    row_numbers = [int(r.strip()) for r in row_numbers_arg.split(",")]

    # Only load this argument if it exists
    zone_name_arg = sys.argv[3] if len(sys.argv) > 3 else None

    wb = xw.Book(PATH_EXCEL)
    sheet = wb.sheets.active
    metadata = get_excel_metadata(sheet)

    column_indices = {header: {"index": idx + 2, "column": col_idx_to_letter(idx + 2)}
                      for idx, header in enumerate(metadata['headers'])}
    formulas_mapping_articles = dynamic_formulas_mapping(FORMULAS_MAPPING_ARTICLES, column_indices)

    if operation == "insert_product_between_columns":
        insert_product_between_columns(sheet=sheet,
                                       row_numbers=row_numbers,
                                       column_indices=column_indices,
                                       formula_mapping=formulas_mapping_articles,
                                       start_column=metadata['headers'][0],
                                       end_column=metadata['headers'][-1])

        # add_or_delete_row_between_columns(sheet=sheet,
        #                                   row_numbers=row_numbers,
        #                                   column_indices=column_indices,
        #                                   start_column=metadata['headers'][0],
        #                                   end_column=metadata['headers'][-1],
        #                                   action="add")

    elif operation == "delete_between_columns":
        add_or_delete_row_between_columns(sheet=sheet,
                                          row_numbers=row_numbers,
                                          column_indices=column_indices,
                                          start_column=metadata['headers'][0],
                                          end_column=metadata['headers'][-1],
                                          action="delete")

    elif operation == "add_zone":
        add_or_delete_row_between_columns(sheet=sheet,
                                          row_numbers=row_numbers,
                                          column_indices=column_indices,
                                          start_column=metadata['headers'][0],
                                          end_column=metadata['headers'][-1],
                                          action="add",
                                          zone="yes",
                                          zone_name=zone_name_arg
                                          )

    else:
        print(f"Unknown operation: {operation}")
