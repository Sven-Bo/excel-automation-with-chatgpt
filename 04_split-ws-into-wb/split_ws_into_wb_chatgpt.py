import xlwings as xw  # Library for interacting with Excel files
from pathlib import Path  # Library for working with file paths

def split_excel_workbook(input_file: str, output_dir: str):
    """Splits an Excel workbook into multiple Excel files, with each file containing one sheet from the original workbook.
    The sheet name is used as the file name for each output file.
    
    Args:
        input_file (str): Path to the input Excel file.
        output_dir (str): Path to the output directory.
    """
    # Set the path to the input Excel file
    excel_file = Path(input_file)

    # Set the path to the output directory
    output_dir = Path(output_dir)

    # Create the output directory if it doesn't exist
    output_dir.mkdir(parents=True, exist_ok=True)

    # Open the input Excel file and create a new hidden Excel app
    with xw.App(visible=False) as app:
        wb = app.books.open(excel_file)

        # Iterate through each sheet in the input workbook
        for sheet in wb.sheets:
            # Create a new workbook
            wb_new = app.books.add()

            # Copy the sheet from the input workbook to the new workbook
            sheet.copy(after=wb_new.sheets[0])

            # Delete the default sheet in the new workbook
            wb_new.sheets[0].delete()

            # Save the new workbook with the sheet name as the file name
            wb_new.save(output_dir / f'{sheet.name}.xlsx')

            # Close the new workbook
            wb_new.close()

# Example usage
split_excel_workbook('Financial_Data.xlsx', 'Output')
