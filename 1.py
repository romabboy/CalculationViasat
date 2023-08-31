import openpyxl
from calculation import calculat, createWB
import os
MAINPATH = r"C:\Users\Dell\Desktop\Робота"

LDir = os.listdir(MAINPATH)
path = filter(lambda p: p.find("My") != -1 and p.find("~") == -1, LDir)
path = list(path)[0]



wb = openpyxl.load_workbook(fr"{MAINPATH}\{path}")
sheet = list(wb)[0]

arr = []
for x in sheet.values:
    arr.append(list(x))
arr = arr[1:]
arr.sort(key=lambda a: a[10])

arr = [[x[1],x[2],x[4],x[9],x[10],x[12],x[13],x[14],x[16],x[17]] for x in arr]



for x in arr:
    x.append(calculat(packet=x[5],date=x[1]))
    if x[8] == "Не оплачено":
        x[-1] = 0

print(arr)
createWB(arr)

print("Everything went well")