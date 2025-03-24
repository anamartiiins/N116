import re

# Function to convert column index to Excel column letter
def col_idx_to_letter(col_idx):
    col_letter = ""
    while col_idx > 0:
        col_idx, remainder = divmod(col_idx - 1, 26)
        col_letter = chr(65 + remainder) + col_letter
    return col_letter


# Function to dynamically replace column names with corresponding letters
def dynamic_formulas_mapping(mapping: dict, column_indices: dict):
    dynamic_mapping = {}

    for key, formula in mapping.items():
        # For each column name in the formula, replace with the corresponding column letter
        for column_name, column_info in column_indices.items():
            column_letter = column_info["column"]
            # Use a regex to ensure we are replacing the exact column name
            formula = re.sub(r'\b' + re.escape(column_name) + r'\b', column_letter, formula)
        dynamic_mapping[key] = formula

    return dynamic_mapping

undo_stack=[]
def undo_last_inserts(sheet):
    global undo_stack

    if not undo_stack:
        print("Nothing to undo.")
        return

    while undo_stack:
        action, row = undo_stack.pop()
        if action == "delete_row":
            sheet.range(f"{row}:{row}").api.Delete()
            print(f"Undo: Deleted row {row}")

