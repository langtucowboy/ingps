import PyGetWindow
import time
import os

z1 = pygetwindow.getAllTitles()
time.sleep(1)
print(len(z1))
# add some app or folder name here
os.startfile("C:\\Users\\NAME\\PATH")
time.sleep(1)
z2 = pygetwindow.getAllTitles()
print(len(z2))
time.sleep(1)
# identifies new window
z3 = [x for x in z2 if x not in z1]
z3 = ''.join(z3)
print(z3)
time.sleep(3)
# selected window manipulation
y = pygetwindow.getWindowsWithTitle(z3)[0]
y.resizeTo(450, 750)
y.moveTo(800, 0)
time.sleep(3)
y.activate()
y.resizeTo(750, 650)