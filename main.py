import pyautogui
from openpyxl import load_workbook
from selenium import webdriver
from webdrivermanager import ChromeDriverManager
import requests
import json
#dia chi lich su hanh trinh cua dinh vi
url = 'http://api.eup.net.vn:8000/ctynguyenngoc/history'
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
    workbook.save(add1)
    # like the doc says, provide API key in header
    API_KEY = '.... your API key ....'
    username = 'ctynguyenngoc'
    password = 'sEj1oRXN0tLOgYJPvRMH'

    session = requests.Session()
    # these are sent along for all requests
    session.headers['X-IG-API-KEY'] = 'af73b810-28d8-4558-9600-02f368221e56'
    # not strictly needed, but the documentation recommends it.
    session.headers['Accept'] = "application/json; charset=UTF-8"

    # log in first, to get the tokens
    response = session.post(
        url + '/session',
        json={'identifier': username, 'password': password},
        headers={'VERSION': '2'},
    )
    print(response)




