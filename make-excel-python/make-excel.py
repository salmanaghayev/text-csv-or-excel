import re
import argparse
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

# Define pattern-replacement rules
replacements = [
    (r'\s{2,}', ','),     # Replace 2+ spaces with comma
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

    # Split each line into a list of fields
    rows = [line.split(',') for line in lines]

    # Load existing workbook and create new sheet
    wb = load_workbook(excel_file)
    if sheet_name in wb.sheetnames:
        print(f"Sheet '{sheet_name}' already exists. Overwriting.")
        del wb[sheet_name]
    ws = wb.create_sheet(title=sheet_name)

    # Write data to the new sheet
    for row_idx, row in enumerate(rows, start=1):
        for col_idx, value in enumerate(row, start=1):
            ws.cell(row=row_idx, column=col_idx, value=value.strip())

    # Save workbook
    wb.save(excel_file)
    print(f"Data written to sheet '{sheet_name}' in '{excel_file}'.")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Clean and write text data to a new sheet in Excel.")
    parser.add_argument("-i", "--input", default="input.txt", help="Input text file")
    parser.add_argument("-e", "--excel", required=True, help="Path to existing Excel file (.xlsx)")
    parser.add_argument("-s", "--sheet", default="ProcessedData", help="Name of new sheet to create")
    args = parser.parse_args()

    main(args.input, args.excel, args.sheet)

# Usage
#python3 convert_to_excel_sheet.py -i input.txt -e existing_file.xlsx -s NewSheetName

