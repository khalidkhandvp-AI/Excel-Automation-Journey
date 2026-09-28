from openpyxl import load_workbook
from openpyxl.styles import Font
wb=load_workbook("Excel Files/Student File.xlsx")
sheet=wb['Sheet1']
cell=sheet.cell(row=21,column=1)
#BOLD THE VALUE
Font_a1=Font(bold=True) 
cell.font=Font_a1
#UNDERLINE THE VALUE
Font_a2=Font(underline="single") 
cell.font=Font_a2
wb.save("Excel Files/Student File.xlsx")