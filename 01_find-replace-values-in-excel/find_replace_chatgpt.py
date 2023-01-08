import os
import openpyxl

# Set the input and output directories
input_dir = 'Input'
output_dir = 'Output'

# Create the output directory if it does not exist
if not os.path.exists(output_dir):
    os.makedirs(output_dir)

# Iterate through all the Excel files in the input directory
for file in os.listdir(input_dir):
    # Check if the file is an Excel file
    if file.endswith('.xlsx'):
        # Open the workbook
        wb = openpyxl.load_workbook(os.path.join(input_dir, file))

        # Iterate through all the sheets in the workbook
        for sheet in wb.worksheets:
            # Iterate through all the cells in the sheet
            for row in sheet.rows:
                for cell in row:
                    # Replace the text if it matches
                    if cell.value == 'Small Business':
                        cell.value = 'Small Market'
                    elif cell.value == 'Midmarket':
                        cell.value = 'Midsize Market'

        # Save the modified workbook to the output directory
        wb.save(os.path.join(output_dir, file))
