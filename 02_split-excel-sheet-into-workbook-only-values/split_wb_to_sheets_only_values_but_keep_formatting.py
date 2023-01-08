# 👉 Save each Excel sheet as separate file. Copy only the values (keep the formatting!)
from pathlib import Path

import xlwings as xw

BASE_DIR = Path(__file__).parent if "__file__" in locals() else Path.cwd()
EXCEL_FILE = BASE_DIR / "data.xlsx"
OUTPUT_DIR = BASE_DIR / "Output"

# Create Output directory
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

with xw.App(visible=False) as app:
    wb = app.books.open(EXCEL_FILE)
    for sheet in wb.sheets:
        # Create a new workbook
        wb_new = app.books.add()

        # Copy the orginal sheet
        sheet.copy(after=wb_new.sheets[0])

        # Clear only the contents
        wb_new.sheets[1].used_range.clear_contents()

        # Delete the inital first sheet when creating the wb
        wb_new.sheets[0].delete()

        # Get the address from the used_range object
        rng_address = sheet.used_range.get_address()

        # Transfer values within used range to new workbook
        wb_new.sheets[0].range(rng_address).value = sheet.range(rng_address).value

        # Save & close workbook
        wb_new.save(OUTPUT_DIR / f'{sheet.name}.xlsx')
        wb_new.close()