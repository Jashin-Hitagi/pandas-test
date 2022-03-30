import os
import tkinter as tk
from tkinter import filedialog

import pandas as pd

'''
一个表里面有多个sheet,在其他几个sheet里面的第一列前面加入第一个sheet的里面指定列
'''

# 打开选择文件夹对话框
print("选择文件夹")
root = tk.Tk()
root.withdraw()
folderPath = filedialog.askdirectory()  # 获得选择好的文件夹
fileSet = os.listdir(folderPath)
print("输入表头，英文逗号分隔")
headerList = input().split(',')
headerMap = {}
for file in fileSet:
    engine = None
    if file.endswith('.xlsx'):
        engine = 'openpyxl'
    if file.endswith('.xls'):
        engine = 'xlrd'
    conv = dict(zip(headerList, [str] * len(headerList)))
    data = pd.read_excel(folderPath + "\\" + file, sheet_name=None, engine=engine, converters=conv)
    for key in data.keys():
        df = data[key]
        if len(headerMap) < 1:
            for header in headerList:
                if df[header].isnull().all():
                    continue
                else:
                    headerMap.update({header: df[header]})
        else:
            for headerKey in headerMap.keys():
                if df.shape[0] < 1:
                    df.insert(0, headerKey, "",allow_duplicates=True)
                elif not df.keys().__contains__(headerKey):
                    df.insert(0, headerKey, headerMap.get(headerKey), allow_duplicates=True)
                elif df.keys().__contains__(headerKey) and df[headerKey].isnull().all():
                    df[headerKey] = headerMap.get(headerKey)

    writer = pd.ExcelWriter(folderPath + "\\" + file)
    for key in data.keys():
        data[key].to_excel(writer, sheet_name=key, index=False)
    writer.save()

print("处理完成,回车键关闭!")
input()
