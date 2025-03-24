from .process_utils import undo_stack

def insert_product_between_columns(sheet, row_numbers: list, column_indices: dict, formula_mapping: dict,
                                   start_column: str,
                                   end_column: str):
    """
    Inserts a blank row below each specified row, only between the given columns.
    Preserves formatting from above and ensures the inserted row has no fill color.
    """
    try:
        global undo_stack

        start_col_idx = column_indices.get(start_column, {}).get("column")
        end_col_idx = column_indices.get(end_column, {}).get("column")

        if start_col_idx is None or end_col_idx is None:
            print(f"Invalid columns: {start_column} or {end_column} not found in column indices.")
            return

        if not isinstance(row_numbers, list):
            row_numbers = [row_numbers]

        for row_number in sorted(row_numbers, reverse=True):
            # Insert row (formats from above are preserved automatically)
            target_range = sheet.range(f"{start_col_idx}{row_number}:{end_col_idx}{row_number}")
            target_range.api.Insert(Shift=2)

            # Get the range of the newly inserted row
            new_range = sheet.range(f"{start_col_idx}{row_number}:{end_col_idx}{row_number}")

            # Clear fill color from each cell in the new row
            for cell in new_range:
                cell.api.Interior.ColorIndex = -4142  # No Fill

            # Apply formulas for specific columns
            for col_name, formula in formula_mapping.items():
                if col_name in column_indices:
                    col_idx = column_indices[col_name]["index"]
                    new_row_cell = sheet.cells(row_number, col_idx)
                    new_row_cell.formula = formula.format(row=row_number)

            undo_stack.append(("delete_row", row_number))
            print(f"Blank row inserted at row {row_number} between columns {start_column} and {end_column}.")

    except Exception as e:
        print(f"Error occurred while inserting the row(s): {e}")



def add_or_delete_row_between_columns(sheet, row_numbers, column_indices: dict, start_column: str, end_column: str,
                                      action: str, zone: str = None, zone_name: str = None):
    """
    Adds or deletes rows within the range between the specified columns.
    Optionally fills the new row with color #F2F2F2 if 'zone' is specified and action is 'add'.
    """

    try:
        global undo_stack

        # Accept a single integer or a list of row numbers
        if not isinstance(row_numbers, list):
            row_numbers = [row_numbers]

        # Sort the row numbers:
        # - descending when adding rows (to avoid shifting)
        # - ascending when deleting rows (to avoid skipping)
        reverse = True if action == "add" else False
        row_numbers = sorted(row_numbers, reverse=reverse)

        start_col_idx = column_indices.get(start_column, {}).get("column")
        end_col_idx = column_indices.get(end_column, {}).get("column")

        if start_col_idx is None or end_col_idx is None:
            print(f"Invalid columns: {start_column} or {end_column} not found in column indices.")
            return

        for row_number in row_numbers:
            target_range = sheet.range(f"{start_col_idx}{row_number}:{end_col_idx}{row_number}")

            if action == "delete":
                target_range.api.Delete(Shift=2)  # Shift cells up
                undo_stack.append(("insert_row", row_number))  # Register undo action
                print(f"Row {row_number} deleted between columns {start_column} and {end_column}.")

            elif action == "add":
                target_range.api.Insert(Shift=2)  # Shift cells down
                if zone == "yes":
                    sheet.cells(row_number+1, column_indices.get("Artigo", {}).get("index")).value = zone_name
                    target_range.api.Interior.Color = 0xF2F2F2  # Light gray fill
                else:
                    new_range = sheet.range(f"{start_col_idx}{row_number}:{end_col_idx}{row_number}")
                    new_range.api.Interior.Color = -4142  # No fill on the new row

                undo_stack.append(("delete_row", row_number))  # Register undo action
                print(f"Row {row_number} added between columns {start_column} and {end_column}.")

            else:
                print(f"Invalid action specified: {action}. Please use 'add' or 'delete'.")

    except Exception as e:
        print(f"Error occurred while modifying the row(s): {e}")


