# first import load_workbook if you have alread an existing excel file
from openpyxl import load_workbook

# creat var and use load_workbook and write the file path for your excel file
workbook = load_workbook(r'C:\Users\LENOVO\Downloads\Analysis\NoofClick.xlsx')
sheet = workbook.active

#start to work
sheet["a1"].value = "Month"

sheet["b1"].value = "Link Clicks"
print(sheet["a1"].value)
print(sheet["b1"].value)

workbook.save(r"C:\Users\LENOVO\Downloads\Analysis\NoofClick.xlsx")