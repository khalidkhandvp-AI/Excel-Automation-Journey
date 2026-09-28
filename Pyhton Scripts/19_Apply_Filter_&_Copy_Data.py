from openpyxl import load_workbook
import pandas as pd
wb=load_workbook("Excel Files/Student File.xlsx")
ws=wb["Sheet1"]
df=pd.read_excel("Excel Files/Student File.xlsx")
filter_data=df[df["Section"]=="Green"]
filter_data.to_excel("Green.xlsx",index=False)
wb.save