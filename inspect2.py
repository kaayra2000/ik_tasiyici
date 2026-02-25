import openpyxl

wb = openpyxl.load_workbook("docs/EN YENİ DK format1 v2.xlsx")
ws = wb.active

for col in range(1, 15):
    print(f"Col {col}: {ws.cell(row=10, column=col).value}")
