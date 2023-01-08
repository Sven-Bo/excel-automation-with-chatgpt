from pathlib import Path

import xlwings as xw  # pip install xlwings

this_dir = Path(__file__).parent if '__file__' in locals() else Path.cwd()
excel_file = this_dir  / 'Financial_Data.xlsx'
output_dir = this_dir  / 'Output'

# Create Output directory
output_dir.mkdir(parents=True, exist_ok=True)

with xw.App(visible=False) as app:
    wb = app.books.open(excel_file)
    for sheet in wb.sheets:
        wb_new = app.books.add()
        sheet.copy(after=wb_new.sheets[0])
        wb_new.sheets[0].delete()
        wb_new.save(output_dir / f'{sheet.name}.xlsx')
        wb_new.close()