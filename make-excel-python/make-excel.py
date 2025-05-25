import re
import argparse
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, PatternFill
from openpyxl.formatting import Rule
from openpyxl.styles.differential import DifferentialStyle
import os

# Define pattern-replacement rules
replacements = [
    (r'\s{2,}', ';'),     # Replace 2+ spaces with semicolon
    (r'\t+', ';'),        # Replace tabs with semicolon
    (r'\s+\|\s+', '|'),   # Replace space-pipe-space with pipe
    (r':', '=')           # Replace colon with equals
]

def process_line(line):
    line = line.strip('"')  # Remove leading/trailing quotes
    for pattern, replacement in replacements:
        line = re.sub(pattern, replacement, line)
    return line

def main(input_file, excel_file, sheet_name):
    # Read and clean lines
    with open(input_file, 'r', encoding='utf-8') as infile:
        lines = [process_line(line) for line in infile]

    # Split each line into a list of fields using semicolon
    rows = [line.split(';') for line in lines]

    # Load or create workbook
    if os.path.exists(excel_file):
        wb = load_workbook(excel_file)
    else:
        from openpyxl import Workbook
        wb = Workbook()
        default_sheet = wb.active
        wb.remove(default_sheet)

    # Delete existing sheet if present
    if sheet_name in wb.sheetnames:
        print(f"Sheet '{sheet_name}' already exists. Overwriting.")
        del wb[sheet_name]

    # Create new sheet
    ws = wb.create_sheet(title=sheet_name)

    # Define header style
    header_fill = PatternFill(start_color="FFD700", end_color="FFD700", fill_type="solid")
    header_font = Font(bold=True, color="000000", size=12)

    # Write data to the new sheet
    for row_idx, row in enumerate(rows, start=1):
        for col_idx, value in enumerate(row, start=1):
            cell = ws.cell(row=row_idx, column=col_idx, value=value.strip())
            if row_idx == 1:  # Apply style to header
                cell.fill = header_fill
                cell.font = header_font

    # Add filter to the headers
    ws.auto_filter.ref = ws.dimensions

    # Add conditional formatting for entire row based on columns B and C
    green_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    red_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")

    green_dxf = DifferentialStyle(fill=green_fill)
    red_dxf = DifferentialStyle(fill=red_fill)

    max_row = ws.max_row
    max_col = ws.max_column
    last_col_letter = get_column_letter(max_col)

    # Green row if both B and C are "up"
    formula_green = "=AND(EXACT($B2,\"up\"), EXACT($C2,\"up\"))"
    ws.conditional_formatting.add(f"A2:{last_col_letter}{max_row}",
        Rule(type="expression", formula=[formula_green], dxf=green_dxf))

    # Red row if any of B or C is not "up"
    formula_red = "=NOT(AND(EXACT($B2,\"up\"), EXACT($C2,\"up\")))"
    ws.conditional_formatting.add(f"A2:{last_col_letter}{max_row}",
        Rule(type="expression", formula=[formula_red], dxf=red_dxf))

    # Save workbook
    try:
        wb.save(excel_file)
        print(f"Data written to sheet '{sheet_name}' in '{excel_file}'.")
    except PermissionError as e:
        print(f"❌ Permission error saving file: {e}")
    except Exception as e:
        print(f"❌ Unexpected error: {e}")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Clean and write text data to a new sheet in Excel.")
    parser.add_argument("-i", "--input", default="input.txt", help="Input text file")
    parser.add_argument("-e", "--excel", required=True, help="Path to existing Excel file (.xlsx)")
    parser.add_argument("-s", "--sheet", default="ProcessedData", help="Name of new sheet to create")
    args = parser.parse_args()

    main(args.input, args.excel, args.sheet)
