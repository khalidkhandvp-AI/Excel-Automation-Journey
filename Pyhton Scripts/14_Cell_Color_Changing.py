from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from openpyxl.styles import Font
wb=load_workbook("Excel Files/Student File.xlsx")
#CHANGING BACKGROUND COLOUR
ws=wb["Sheet1"]
ws["d5"].fill=PatternFill(fill_type="solid",fgColor="0000FF")
wb.save("Excel Files/Student File.xlsx")
#CHANGING THE TEXT COLOUR
ws=wb["Sheet1"]
ws["d8"].font=PatternFill(fill_type="solid",fgColor="FF0000")
wb.save("Excel Files/Student File.xlsx")
