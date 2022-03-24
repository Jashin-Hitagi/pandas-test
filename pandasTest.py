import numpy as np
import pandas as pd
import tkinter as tk
from tkinter import filedialog
import os


# s = pd.Series([1, 3, 5, np.nan, 6, 8])
# print(s)
#
# dates = pd.date_range('20130101', periods=6)
# print(dates)

# df = pd.DataFrame(np.random.randn(6, 4), index=dates, columns=list('ABCD'))
# print(df.index)
# print(df)

# df2 = pd.DataFrame({'A': 1.,
# 'B': pd.Timestamp('20130102'),
# 'C': pd.Series(1, index=list(range(4)), dtype='float32'),
# 'D': np.array([3] * 4, dtype='int32'),
# 'E': pd.Categorical(["test", "train", "test", "train"]),
# 'F': 'foo'})
#
# print(df2)
# print(df2.dtypes)

'''打开选择文件夹对话框'''
root = tk.Tk()
root.withdraw()
folderPath = filedialog.askdirectory()  # 获得选择好的文件夹
fileSet = os.listdir(folderPath)
fileName = str('/result.xlsx')
file_path = folderPath + fileName
if os.path.exists(file_path):
    os.remove(file_path)
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
# writer.close()

