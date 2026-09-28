from openpyxl import load_workbook
#SHEET NAME 
student_data=load_workbook("Excel Files/Student File.xlsx")
print(student_data.sheetnames)
#D7 CELL DATA
Sheet1=student_data.active
value_d7=Sheet1.cell(row=5,column=4).value
print(value_d7)
#C8 CELL DATA
value_c8=Sheet1.cell(row=5,column=3).value
print(value_c8)