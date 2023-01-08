from pathlib import Path
import openpyxl

BASE_DIR = Path(__file__).parent
INPUT_DIR = BASE_DIR / "Input"
OUTPUT_DIR = BASE_DIR / "Output"

# Create output directory
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

replacement_pair = {"Small Business": "Small Market", "Midmarket": "Midsize Market"}

files = list(INPUT_DIR.rglob("*.xls*"))
for file in files:
    wb = openpyxl.load_workbook(file)
    for ws in wb.worksheets:
        # Iterate over the columns and rows, search for the text and replace
        for row in ws.iter_rows():
            for cell in row:
                if cell.value in replacement_pair.keys():
                    cell.value = replacement_pair.get(cell.value)
    wb.save(OUTPUT_DIR / f"{file.stem}_NEW.xlsx")