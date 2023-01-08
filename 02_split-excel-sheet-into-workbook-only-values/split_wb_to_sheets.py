import os
import openpyxl

# open workbook
wb = openpyxl.load_workbook('data.xlsx')

# create 'Output' folder if it doesn't exist
if not os.path.exists('Output'):
    os.makedirs('Output')

# save each worksheet as a new workbook
for sheet in wb:
    wb_temp = openpyxl.Workbook()
    ws_temp = wb_temp.active
    for row in sheet:
        for cell in row:
            ws_temp[cell.coordinate].value = cell.value
            ws_temp[cell.coordinate].font = cell.font
            ws_temp[cell.coordinate].border = cell.border
            ws_temp[cell.coordinate].fill = cell.fill
            ws_temp[cell.coordinate].number_format = cell.number_format
            ws_temp[cell.coordinate].protection = cell.protection
            ws_temp[cell.coordinate].alignment = cell.alignment
    wb_temp.save(f'Output/{sheet.title}.xlsx')

# close workbook
wb.close()
