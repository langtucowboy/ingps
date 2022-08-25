import pyautogui
from openpyxl import load_workbook
from selenium import webdriver
from webdrivermanager import ChromeDriverManager
import pygetwindow as gw
import os
import time
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
for i in range(2,3):
    soxe = datasheet.cell(i,2).value
    shipment = datasheet.cell(i,1).value
    chitiet["B5"] = soxe
    chitiet["B7"] = shipment
    workbook.save(add1)
    a = gw.getWindowsWithTitle('19136 - CÔNG TY CỔ PHẦN NGUYỄN NGỌC LOGISTICS - v1.0.15.217 - Google Chrome')[0]
    a.activate()
    pyautogui.moveTo(1234,202)
    pyautogui.leftClick()
    pyautogui.write(soxe)
    pyautogui.moveTo(1186, 264)
    pyautogui.leftClick()
    pyautogui.moveTo(1114, 354)
    pyautogui.leftClick()
    pyautogui.moveTo(1343, 438, duration =1)
    pyautogui.leftClick()
    pyautogui.moveTo(1250, 260, duration= 2)
    pyautogui.scroll(-300)
    pyautogui.moveTo(1121, 367, duration=1)
    pyautogui.leftClick()
    pyautogui.moveTo(94, 693, duration=1)
    pyautogui.leftClick()
