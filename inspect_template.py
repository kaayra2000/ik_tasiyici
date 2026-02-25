import openpyxl

wb = openpyxl.load_workbook("docs/EN YENİ DK format1 v2.xlsx")
ws = wb.active

for row in ws.iter_rows():
    for cell in row:
        if cell.value in ["o", "h", "e", "Proje", "e ", "h "]:
            print(f"Cell {cell.coordinate}: {repr(cell.value)}")
