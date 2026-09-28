from openpyxl import load_workbook
wb=load_workbook("Excel Files/Student File.xlsx")
ws=wb["Sheet2"]
ws.delete_cols(1)
wb.save("Excel Files/Student File.xlsx")