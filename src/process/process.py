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


def create_supplier_sheets(sheet: xw.Sheet, only_supplier: str):
    """
    Create one sheet per supplier by copying the 'Fornecedor_Template' sheet,
    filtering rows and clearing unrelated production/material values, and
    copying only the images of the selected articles.

    sheet (xw.Sheet): The source sheet containing the full budget table.
    only_supplier (str): If provided, generate only this supplier.

    Returns:
        None. New sheets are added to the workbook.
    """

    ## IDENTIFY TEMPLATE
    template = sheet.book.sheets["Fornecedor_Template"]

    ## IDENTIFY COLUMNS THAT NEED TO BE COPIED TO FORNECEDOR SHEET
    supplier_columns = [
        "Fornecedor Produção 1", "Fornecedor Produção 2",
        "Fornecedor Tecido 1",    "Fornecedor Tecido 2",
        "Fornecedor Tecido 3"
    ]

    supplier_values = [
        "Produção 1", "Produção 2",
        "Material/Tecido 1", "Material/Tecido 2", "Material/Tecido 3"
    ]

    columns_to_keep = [
        "Artigo", "Descrição", "Imagem de Referência", "Qtd", "Dimensões",
        "Acabamentos", "Notas", "Break", "Produção 1", "Produção 2",
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

    ## FIND UNIQUE SUPPLIERS
    unique_suppliers = pd.unique(df[supplier_columns].values.ravel())
    unique_suppliers = [s for s in unique_suppliers if isinstance(s, str) and s.strip()]

    if only_supplier:
        only_supplier_norm = only_supplier.strip().upper()
        unique_suppliers = [
            s for s in unique_suppliers
            if s.strip().upper() == only_supplier_norm
        ]

    ## START ITERATING FROM UNIQUE SUPPLIERS
    for supplier in unique_suppliers:

        # FILTER DATAFRAME GETTING ONLY THE ROWS OF PRODUCTS RELATED WITH THE FORNECEDOR
        df_filtered = df[df[supplier_columns].apply(lambda row: supplier in row.values, axis=1)].copy()

        if df_filtered.empty:
            continue

        df_filtered = df_filtered.loc[:, columns_to_keep].reset_index(drop=True)

        # CREATE THE SHEET FOR EACH FORNECEDOR
        if supplier in [sht.name for sht in sheet.book.sheets]:
            sheet.book.sheets[supplier].delete()

        new_sheet = template.copy(name=supplier, after=template)
        new_sheet.range("B3").value = supplier

        # ADJUST RANGE AND COPY ALL DF VALUES
        target_range = new_sheet.range("B7").resize(df_filtered.shape[0], df_filtered.shape[1])
        target_range.value = df_filtered.values.tolist()

        # IDENTIFY SUPPLIER SHEET HEADERS
        header_row_supplier = 6
        headers_supplier = new_sheet.range((header_row_supplier, 1)).expand("right").value
        header_columns_index_supplier = {str(h).strip(): idx+1 for idx, h in enumerate(headers_supplier) if h}

        # LOOP THROUGH EACH PASTED ROW
        for i_row in range(header_row_supplier + 1, header_row_supplier + 1 + df_filtered.shape[0]):
            # FOR EACH SUPPLIER-COLUMN / VALUE-COLUMN PAIR
            for j, sup_col in enumerate(supplier_columns):
                # READ THE SUPPLIER NAME IN THIS ROW/COLUMN
                cell_supplier = new_sheet.cells(i_row, header_columns_index_supplier[sup_col]).value
                # IF IT'S NOT THE SHEET'S SUPPLIER, BLANK BOTH
                if str(cell_supplier).strip().upper() != supplier.strip().upper():
                    # CLEAR THE SUPPLIER NAME CELL
                    new_sheet.cells(i_row, header_columns_index_supplier[sup_col]).value = None
                    # CLEAR THE CORRESPONDING PRODUCTION/MATERIAL VALUE
                    val_col = supplier_values[j]
                    new_sheet.cells(i_row, header_columns_index_supplier[val_col]).value = None

            # INSERT DYNAMIC FORMULAS FOR COSTS
            cost_unit_column = header_columns_index_supplier["Custo Unitário"]
            formula_unit = f"=SUM(J{i_row}:N{i_row})"
            new_sheet.cells(i_row, cost_unit_column).formula = formula_unit

            cost_total_column = header_columns_index_supplier["Custo Total"]
            formula_total = f"=O{i_row}*E{i_row}"
            new_sheet.cells(i_row, cost_total_column).formula = formula_total

        ##TODO: adicionar cópia das imagens. adicionar formatação e remoção de linhas extra.
    print('ana')