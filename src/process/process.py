import pandas as pd
from src.constants import HEADER_START
import xlwings as xw


def insert_product_rows(sheet, row_numbers: list, column_indices: dict, formula_mapping: dict):
    """
    Inserts a full blank row below each specified row and applies formulas in the specified columns.
           sheet (xw.Sheet): Excel sheet object.
           row_numbers (list): List of row numbers to insert after.
           column_indices (dict): Mapping of column names to their index and letter.
           formula_mapping (dict): Formulas to apply by column name.
    Returns:
        None
    """
    try:
        if not isinstance(row_numbers, list):
            row_numbers = [row_numbers]

        row_numbers = [int(r) for r in row_numbers]

        for row_number in sorted(row_numbers, reverse=False):
            # Insert full row (preserves formatting)
            sheet.api.Rows(row_number).Insert()

            # Apply formulas to defined columns
            for col_name, formula in formula_mapping.items():
                if col_name in column_indices:
                    col_idx = column_indices[col_name]["index"]
                    formula_cell = sheet.cells(row_number, col_idx)
                    formula_cell.formula = formula.format(row=row_number)

            print(f"Full row inserted at row {row_number} with formulas applied.")

    except Exception as e:
        print(f"Error occurred while inserting rows: {e}")

def insert_zone_row(sheet, row_numbers: list, column_indices: dict, zone_name: str):
    """
    Adds formatted zone rows at the given row numbers, limited to columns from start to end.
           sheet (xw.Sheet): Excel sheet object.
           row_numbers (list): List of row numbers where zone rows should be inserted.
           column_indices (dict): Dict with column name -> {index, column letter}.
           zone_name (str): Name of the zone to write into the "Artigo" column.
    Returns:
        None
    """
    try:
        if not isinstance(row_numbers, list):
            row_numbers = [row_numbers]

        row_numbers = [int(r) for r in row_numbers]

        start_col_letter = column_indices.get("Artigo", {}).get("column")
        end_col_letter = list(column_indices.values())[-1]["column"]

        if not start_col_letter or not end_col_letter:
            print("Column boundaries not found.")
            return

        for row_number in sorted(row_numbers, reverse=True):
            # Insert new row
            sheet.api.Rows(row_number).Insert()

            # Format the new row between the defined columns
            target_range = sheet.range(f"{start_col_letter}{row_number}:{end_col_letter}{row_number}")
            target_range.api.Interior.Color = 0xF2F2F2
            sheet.api.Rows(row_number).RowHeight = 15

            # Write the zone name under the "Artigo" column and apply bold
            artigo_col_index = column_indices.get("Artigo", {}).get("index")
            if artigo_col_index:
                cell = sheet.cells(row_number, artigo_col_index)
                cell.value = zone_name
                cell.api.Font.Bold = True
            print(f"Zone '{zone_name}' added at row {row_number} from {start_col_letter} to {end_col_letter}")

    except Exception as e:
        print(f"Error adding zone row(s): {e}")


def insert_product_between_columns(sheet, row_numbers: list, column_indices: dict, formula_mapping: dict,
                                   start_column: str,
                                   end_column: str):
    """
    Inserts a blank row below each specified row, only between the given columns.
    Preserves formatting from above and ensures the inserted row has no fill color.
    """
    try:

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
                    sheet.cells(row_number + 1, column_indices.get("Artigo", {}).get("index")).value = zone_name
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


def create_supplier_sheets(sheet: xw.Sheet, metadata: dict, only_supplier: str = None ):
    """
    Duplicate the orçamento sheet for each supplier, filtering only the rows
    where that supplier appears in any of the specified supplier columns.

    Parameters:
    sheet (xw.Sheet): The source sheet to duplicate and filter.
    only_supplier (str, optional): If provided, generate only this supplier's sheet.

    Returns:
        None. New sheets are added to the workbook (or replaced if existing).
    """
    wb = sheet.book

    # Define which columns contain supplier names
    supplier_columns = [
        "Fornecedor Produção 1", "Fornecedor Produção 2",
        "Fornecedor Tecido 1",    "Fornecedor Tecido 2",
        "Fornecedor Tecido 3"
    ]

    supplier_values = [
        "Produção 1", "Produção 2",
        "Material/Tecido 1", "Material/Tecido 2", "Material/Tecido 3"
    ]

    # Keep only specific columns in final sheet in the correct order
    columns_to_keep = [
                          "Artigo", "Descrição", "Imagem de Referência", "Qtd", "Dimensões",
                          "Acabamentos", "Break", "Produção 1", "Produção 2",
                          "Material/Tecido 1", "Material/Tecido 2", "Material/Tecido 3",
                          "Custo Unitário", "Custo Total", "Break2"
                      ] + supplier_columns

    ## READ DATAFRAME FROM ORÇAMENTO SHEET
    header_range = sheet.range(HEADER_START)
    start_row = header_range.row
    start_col = header_range.column

    data_range = sheet.range((start_row, start_col)).expand("table")
    df = data_range.options(pd.DataFrame, header=1, index=False).value
    df = df.loc[:, ~df.columns.astype(str).str.contains("^Unnamed")]

    # Identify unique suppliers
    uniq = pd.unique(df[supplier_columns].values.ravel())
    suppliers = [s for s in uniq if isinstance(s, str) and s.strip()]
    if only_supplier:
        target = only_supplier.strip().upper()
        suppliers = [s for s in suppliers if s.strip().upper() == target]

    orig_start_row, orig_start_col = start_row, start_col

    # For each supplier, copy and filter
    for sup in suppliers:
        # Delete existing sheet if present
        if sup in [sh.name for sh in wb.sheets]:
            wb.sheets[sup].delete()
        # Copy the orçamento sheet
        new_sh = sheet.copy(name=sup, after=sheet)

        # Determine headers and indices
        headers = new_sh.range((start_row, start_col)).expand("right").value
        header_idx = {h: idx + start_col for idx, h in enumerate(headers)}
        last_row = new_sh.range((start_row, start_col)).expand("down").last_cell.row

        tbl = new_sh.range((start_row, start_col)).expand("table")
        tbl.value = tbl.value

        start_row_loop = orig_start_row

        # Delete any rows above the header row
        if start_row_loop > 1:
            new_sh.api.Rows(f"1:{start_row_loop-1}").Delete()
            # Adjust last_row after deletion
            last_row = last_row - (start_row_loop - 1)
            start_row_loop = 1

        # Delete any rows below the last data row
        max_row = new_sh.api.UsedRange.Rows.Count
        if last_row < max_row:
            new_sh.api.Rows(f"{last_row+1}:{max_row}").Delete()

        # Collect rows to delete (skip header row) (skip header row)
        to_delete = []
        for row in range(start_row_loop + 1, last_row + 1):
            row_vals = [
                new_sh.cells(row, header_idx[col]).value for col in supplier_columns
            ]
            if not any(isinstance(v, str) and v.strip().upper() == sup.strip().upper() for v in row_vals):
                to_delete.append(row)

        # Delete all rows at once if any
        if to_delete:
            ranges = ",".join(f"{r}:{r}" for r in to_delete)
            try:
                new_sh.api.Range(ranges).Delete()
            except Exception:
                for r in reversed(to_delete):
                    new_sh.api.Rows(r).Delete()

        # Clear supplier names and cost values not matching current supplier
        for i_row in range(start_row_loop + 1, last_row + 1):
            # For each supplier-column / value-column pair
            for j, sup_col in enumerate(supplier_columns):
                cell_supplier = new_sh.cells(i_row, header_idx[sup_col]).value
                # If it is not the current sheet's supplier, clear both cells
                if not (isinstance(cell_supplier, str) and cell_supplier.strip().upper() == sup.strip().upper()):
                    # Clear supplier name cell
                    new_sh.cells(i_row, header_idx[sup_col]).value = None
                    # Clear corresponding cost value cell
                    val_col = supplier_values[j]
                    new_sh.cells(i_row, header_idx[val_col]).value = None

        # Retrieve current headers
        current_headers = new_sh.range((start_row_loop, start_col)).expand("right").value
        # Delete columns not in columns_to_keep (right to left)
        for idx in range(len(current_headers)-1, -1, -1):
            if current_headers[idx] not in columns_to_keep:
                new_sh.api.Columns(start_col + idx).Delete()

        # Recompute headers and indices
        headers_updated = new_sh.range((start_row_loop, start_col)).expand("right").value
        header_idx = {h: i + start_col for i, h in enumerate(headers_updated)}
        last_row = new_sh.range((start_row_loop, start_col)).expand("down").last_cell.row

        insert_col = start_col + headers_updated.index("Acabamentos") + 1
        new_sh.api.Columns(insert_col).Insert()

        # Set header
        new_sh.cells(start_row_loop, insert_col).value = "Notas"

        # Fill cells based on supplier tissue columns
        for row in range(start_row_loop + 1, last_row + 1):
            v1 = new_sh.cells(row, header_idx.get("Fornecedor Tecido 1")+1).value
            v2 = new_sh.cells(row, header_idx.get("Fornecedor Tecido 2")+1).value
            v3 = new_sh.cells(row, header_idx.get("Fornecedor Tecido 3")+1).value
            if (not v1 and not v2 and not v3):
                new_sh.cells(row, insert_col).value = ""
            else:
                new_sh.cells(row, insert_col).value = "Inserir Fornecedor de Produção correspondente"
                # Enable wrap text for notes cell
                new_sh.cells(row, insert_col).api.WrapText = True
                # Set font color to red
                new_sh.cells(row, insert_col).api.Font.Color = 0x0000FF

        last_row = new_sh.range((start_row_loop, start_col)).expand("down").last_cell.row

        # INSERT DYNAMIC FORMULAS FOR COSTS
        for i_row in range(start_row_loop + 1, last_row + 1):
            cost_unit_column = header_idx["Custo Unitário"]+1
            formula_unit = f"=SUM(J{i_row}:N{i_row})"
            new_sh.cells(i_row, cost_unit_column).formula = formula_unit

            cost_total_column = header_idx["Custo Total"]+1
            formula_total = f"=O{i_row}*E{i_row}"
            new_sh.cells(i_row, cost_total_column).formula = formula_total


        ## TOTALS OF CUSTS
        # Copy rows 1:5 from the Fornecedor_Template sheet
        template = sheet.book.sheets["Fornecedor_Template"]
        num_rows_above = 5
        rows_to_copy = template.range(f"1:{num_rows_above}")

        # Insert the same number of rows above the header in the supplier sheet
        new_sh.api.Rows(f"1:{num_rows_above}").Insert()

        # Paste content and formatting
        rows_to_copy.api.WrapText=True
        rows_to_copy.api.Copy()
        dest_range = new_sh.range((1, 1)).resize(num_rows_above, rows_to_copy.columns.count)
        dest_range.api.PasteSpecial(Paste=-4104)  # xlPasteAll
        dest_range.api.WrapText=True
        dest_range.api.VerticalAlignment = xw.constants.VAlign.xlVAlignCenter

        # Insert cell with the name of the multiplier to allow vlookup to work
        new_sh.cells(3, start_col).value = sup
        new_sh.cells(3, start_col).api.Font.Color = 0xFFFFFF


