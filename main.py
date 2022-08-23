import pyautogui
from openpyxl import load_workbook
#load excel file
workbook = load_workbook(filename="C:\\Users\Work\Desktop\chitietlotrinh.xlsx")
#open workbook
tensheet = workbook.sheetnames
print(tensheet)
#chon sheet1 de in
chitiet = workbook["Sheet1"]
print(chitiet["B6"].value)
def lammoi():
    chitiet["B5"] = ""
    chitiet["B7"] = ""
    chitiet.delete_rows(idx = 9, amount = 300)
    workbook.save(filename="C:\\Users\Work\Desktop\chitietlotrinh.xlsx")
lammoi()
print(chitiet["B11"].value)



