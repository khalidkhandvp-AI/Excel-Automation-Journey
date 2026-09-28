from openpyxl import Workbook
workbook=Workbook()
Sheet1=workbook.active
Sheet1.cell(row=3,column=3,value='Khalid')
workbook.save("created by python.xlsx")
