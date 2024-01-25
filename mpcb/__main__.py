from openpyxl import load_workbook, Workbook
import re

prompt = "varroc"

new_wb = Workbook()
new_sheet = new_wb.active

workbook = load_workbook("../files/testfile.xlsx")

def write_to_file(row, name):
    data = []
    for cell in row:
        if cell.value:
            data.append(cell.value)
    data.append(name)
    new_sheet.append(data)
    pass

for sheet in workbook.worksheets:
    for i, cell in enumerate(sheet["B"]):
        val = cell.value
        try:
            if val and re.search(prompt, val, re.IGNORECASE):
                write_to_file(sheet[i+1], sheet.title)
        except Exception:
            continue

new_wb.save(filename=f"{prompt}.xlsx")
