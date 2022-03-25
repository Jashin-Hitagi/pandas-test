import pandas as pd
import tkinter as tk
from tkinter import filedialog
import os

'''打开选择文件夹对话框'''
root = tk.Tk()
root.withdraw()
folderPath = filedialog.askdirectory()  # 获得选择好的文件夹
fileName = str('/result.xlsx')
file_path = folderPath + fileName
if os.path.exists(file_path):
    os.remove(file_path)
fileSet = os.listdir(folderPath)
sheetMap = {}
for file in fileSet:

    data = pd.read_excel(folderPath + "\\" + file, sheet_name=None)
    if len(sheetMap) < 1:
        for key in data.keys():
            sheetMap.update({key: list()})
    # print(data.values())

    for sheet in sheetMap:
        sheetMap[sheet].append(data.get(sheet))

writer = pd.ExcelWriter(file_path)
for sheet in sheetMap:
    result = pd.concat(sheetMap[sheet])
    result.to_excel(writer, sheet_name=sheet, index=False)
writer.save()
print("合并完成,回车键关闭!")
input()
