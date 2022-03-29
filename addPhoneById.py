import os
import tkinter as tk
from tkinter import filedialog

import pandas as pd
from phone import Phone

'''
没有手机号的sheet要根据ID去另一个sheet查手机号然后添加进去
'''


class phoneInfo(object):

    def __init__(self, personId, areaCode=None, phoneNum=None, address=None):
        self.address = address
        self.phoneNum = phoneNum
        self.areaCode = areaCode
        self.personId = personId


class addressInfo(object):

    def __init__(self, areaCode, address=None):
        self.address = address
        self.areaCode = areaCode


def get_address(ph, countryMap, areaCode, phoneNum):
    if '86' == str(areaCode) or '+86' == str(areaCode):
        area = ph.find(phoneNum)
        return "中国" + area.get('province') + area.get('city')
    else:
        if countryMap.__contains__(areaCode):
            return countryMap.get(areaCode).address


def get_countryCode():
    countryMap = {}
    dat_file = os.path.join(os.path.dirname(__file__), "country.csv")
    countryPhone = pd.read_csv(dat_file)
    for lineIndex in range(countryPhone.shape[0] - 1):
       line = countryPhone.loc[lineIndex]
       countryMap.update({line['phonecode']:addressInfo(line['phonecode'], line['name_zh'])})
    return countryMap



# 打开选择文件夹对话框
root = tk.Tk()
root.withdraw()
folderPath = filedialog.askdirectory()  # 获得选择好的文件夹
countryMap = get_countryCode()
phone_dat_file = os.path.join(os.path.dirname(__file__), "phone.dat")
p = Phone(phone_dat_file)
fileSet = os.listdir(folderPath)
phoneMap = {}
for file in fileSet:
    engine = None
    if file.endswith('.xlsx'):
        engine = 'openpyxl'
    data = pd.read_excel(folderPath + "\\" + file, sheet_name=None, engine=engine)
    for key in data.keys():
        df = data[key]
        if len(phoneMap) < 1:
            for index in range(len(df)):
                a = df.loc[index]
                person = phoneInfo(a["用户ID"], a["区号"], a["手机号"], get_address(p, countryMap, a["区号"], a["手机号"]))
                phoneMap.update({a["用户ID"]: person})

        if not df.keys().__contains__('区号') or len(df['区号']) < 1:
            areaCodeList = []
            for personId in df["用户ID"]:
                areaCodeList.append(phoneMap.get(personId).areaCode)
            df.insert(df.shape[1], '区号', areaCodeList)

        if not df.keys().__contains__('手机号') or len(df['手机号']) < 1:
            areaCodeList = []
            for personId in df["用户ID"]:
                areaCodeList.append(phoneMap.get(personId).phoneNum)
            df.insert(df.shape[1], '手机号', areaCodeList)

        if not df.keys().__contains__('用户手机号归属地') or len(df['用户手机号归属地']) < 1:
            areaCodeList = []
            for personId in df["用户ID"]:
                areaCodeList.append(phoneMap.get(personId).address)
            df.insert(df.shape[1], '用户手机号归属地', areaCodeList)

    writer = pd.ExcelWriter(folderPath + "\\" + file)
    for key in data.keys():
        data[key].to_excel(writer, sheet_name=key, index=False)
    writer.save()

print("处理完成,回车键关闭!")
input()
