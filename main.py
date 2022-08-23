import pyautogui
from openpyxl import load_workbook
from selenium import webdriver
from webdrivermanager import ChromeDriverManager
#dia chi lich su hanh trinh cua dinh vi
url = "http://fms.ctms.vn/#history_path.html"
#dia chi file chi tiet (ghi tat add1)
add1 = "C:\\Users\Work\Desktop\chitietlotrinh.xlsx"
#load excel file
workbook = load_workbook(filename=add1)
#open workbook
tensheet = workbook.sheetnames
print(tensheet)
#goi ten sheet
chitiet = workbook["Sheet1"]
datasheet = workbook["Data"]
def lammoi():
    chitiet["B5"] = ""
    chitiet["B7"] = ""
    chitiet.delete_rows(idx=9, amount=1000)
    workbook.save(add1)
#dem so dong o sheet data
mr = datasheet.max_row
#bat dau tung dong cua file data
for i in range(2,mr+1):
    soxe = datasheet.cell(i,2).value
    shipment = datasheet.cell(i,1).value
    chitiet["B5"] = soxe
    chitiet["B7"] = shipment
    browser = webdriver.Chrome('C:\\Users\Work\PycharmProjects\helloworld\venv\Scripts\chromedriver.exe')
    browser.get(url)




