import re

# Function to convert column index to Excel column letter
def col_idx_to_letter(col_idx):
    """
    Converts a column index (1-based) to its corresponding Excel column letter.
        col_idx (int): Excel column index (e.g., 1 → 'A').
    Returns:
        str: Excel-style column letter.
    """
    col_letter = ""
    while col_idx > 0:
        col_idx, remainder = divmod(col_idx - 1, 26)
        col_letter = chr(65 + remainder) + col_letter
    return col_letter


# Function to dynamically replace column names with corresponding letters
def dynamic_formulas_mapping(mapping: dict, column_indices: dict):
    """
    Replaces column names in formulas with their corresponding Excel letters.
        mapping (dict): Dictionary of formulas with column names as placeholders.
        column_indices (dict): Mapping of column names to their index and Excel letter.
    Returns:
        dict: Updated formulas with column names replaced by Excel column letters.
    """
    dynamic_mapping = {}

    for key, formula in mapping.items():
        # For each column name in the formula, replace with the corresponding column letter
        for column_name, column_info in column_indices.items():
            column_letter = column_info["column"]
            # Use a regex to ensure we are replacing the exact column name
            formula = re.sub(r'\b' + re.escape(column_name) + r'\b', column_letter, formula)
        dynamic_mapping[key] = formula

    return dynamic_mapping


def find_cell_by_content(sheet, content, return_type="ref"):
    """
    Finds the first cell that matches the given content and returns the desired reference format.
           sheet (xw.Sheet): Excel sheet object.
           content (str): Content to search for.
           return_type (str): Desired return format: 'ref', 'cell', 'letter', 'index', or 'row'.
    Returns:
        str, int or xw.Range: Reference or cell info, depending on return_type.
    """
    for row in sheet.used_range.rows:
        for cell in row:
            if cell.value == content:
                col_letter = col_idx_to_letter(col_idx=cell.column)
                if return_type == "cell":
                    return cell
                elif return_type == "letter":
                    return col_letter
                elif return_type == "index":
                    return cell.column
                elif return_type == "row":
                    return cell.row
                else:  # default to Excel-style reference
                    return f"{col_letter}{cell.row}"
    return None




##FIXME: validar se pode ficar a opção de eliminar como fazem atualmente.
# undo_stack=[]
# def undo_last_inserts(sheet):
#     global undo_stack
#
#     if not undo_stack:
#         print("Nothing to undo.")
#         return
#
#     while undo_stack:
#         action, row = undo_stack.pop()
#         if action == "delete_row":
#             sheet.range(f"{row}:{row}").api.Delete()
#             print(f"Undo: Deleted row {row}")

