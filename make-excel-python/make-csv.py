import re
import argparse

# Define your pattern-replacement pairs here
replacements = [
    (r'\s{2,}', ','),     # Replace 2+ spaces with comma
    (r'\t+', ';'),        # Replace tabs with semicolon
    (r'\s+\|\s+', '|'),   # Replace space-pipe-space with pipe
    (r':', '=')           # Replace colon with equals
]

def process_line(line):
    # Remove leading and trailing quotes
    line = line.strip('"')

    # Apply each pattern replacement
    for pattern, replacement in replacements:
        line = re.sub(pattern, replacement, line)
    
    return line

def main(input_file, output_file):
    with open(input_file, 'r', encoding='utf-8') as infile, \
         open(output_file, 'w', encoding='utf-8') as outfile:
        
        for line in infile:
            cleaned_line = process_line(line)
            outfile.write(cleaned_line + '\n')

    print(f"CSV file created at: {output_file}")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Clean and convert text to CSV with custom rules.")
    parser.add_argument("-i", "--input", default="input.txt", help="Input text file")
    parser.add_argument("-o", "--output", default="output.csv", help="Output CSV file")
    args = parser.parse_args()

    main(args.input, args.output)

# Usage
# .\Convert-To-CustomCSV.ps1 -InputFile "mydata.txt" -OutputFile "cleaned_output.csv"
