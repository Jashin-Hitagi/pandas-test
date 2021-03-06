import os
import tkinter as tk
from tkinter import filedialog

import pandas as pd

'''
没有targetCol的sheet要根据sheetId去另一个sheet查targetCol然后添加进去
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


# 归属地查询
def get_address(ph, countryMap, areaCode, phoneNum):
    if '86' == str(areaCode) or '+86' == str(areaCode):
        pd = ph[ph['号段'].isin([phoneNum[:7]])]
        if pd['省区'].isnull().all():
            return ""
        else:
            province = pd.loc[list(pd.index)[0], '省区']
            city = pd.loc[list(pd.index)[0], '城市']
            if province == city:
                return "中国" + province
            else:
                return "中国" + province + city
    else:
        if countryMap.__contains__(areaCode):
            return countryMap.get(areaCode).address


def get_countryCode():
    countryMap = {}
    dat_file = os.path.join(os.path.dirname(__file__), "country.csv")
    countryPhone = pd.read_csv(dat_file)
    for lineIndex in range(countryPhone.shape[0] - 1):
        line = countryPhone.loc[lineIndex]
        countryMap.update({line['phonecode']: addressInfo(line['phonecode'], line['name_zh'])})
    return countryMap


def get_chineseCode():
    dat_file = os.path.join(os.path.dirname(__file__), "phone481520.csv")
    return pd.read_csv(dat_file,converters={'号段':str})


# 打开选择文件夹对话框
root = tk.Tk()
root.withdraw()
folderPath = filedialog.askdirectory()  # 获得选择好的文件夹
print("输入需要匹配的表头：")
sheetId = input().strip()
print("是否添加手机号归属地(y/n): ")
localFlag = input().strip()
phoneHeader = None
countryMap = {}
chineseMap = pd.DataFrame
if localFlag == 'y':
    print("输入区号表头和手机号表头，逗号分隔: ")
    phoneHeader = input().replace("，",',').split(',')
print("输入除归属地以外需要插入的表头，逗号分隔: ")
headerList = input().strip().replace("，",',').split(',')
print("开始运行... ")
if localFlag == 'y':
    countryMap = get_countryCode()
    chineseMap = get_chineseCode()
fileSet = os.listdir(folderPath)
infoMap = {}
for file in fileSet:
    engine = None
    if not file.endswith('.xlsx') and not file.endswith('.xls'):
        continue
    if file.endswith('.xls'):
        engine = 'xlrd'
    if file.endswith('.xlsx'):
        engine = 'openpyxl'
    conv = dict(zip(headerList, [str] * len(headerList)))
    conv.update({sheetId:str})
    data = pd.read_excel(folderPath + "\\" + file, sheet_name=None, engine=engine, converters=conv)
    for key in data.keys():
        df = data[key]
        if not df.keys().__contains__(sheetId) or df[sheetId].isnull().all():
            continue

        if len(infoMap) < 1:
            for index in range(len(df)):
                infoList = []
                a = df.loc[index]
                if a[sheetId] is None:
                    continue
                for header in headerList:
                    infoList.append((header, a[header]))
                if localFlag == 'y' and len(phoneHeader) > 1:
                    infoList.append(
                        ('用户手机号归属地', get_address(chineseMap, countryMap, a[phoneHeader[0]], a[phoneHeader[1]])))
                infoMap.update({str(a[sheetId]).replace(" ",""): infoList})

        for header in headerList:
            if not df.keys().__contains__(header) or df[header].isnull().all():
                colList = []
                for personId in df[sheetId]:
                    if personId is None:
                        personId = str(personId).replace(" ","")
                        colList.append("")
                    elif infoMap.get(personId) is not None:
                        for info in infoMap.get(personId):
                            if info[0] == header:
                                colList.append(info[1])
                    else:
                        colList.append("")
                df.loc[:, header] = colList

        if localFlag == 'y' and len(phoneHeader) > 1:
            colList = []
            for personId in df[sheetId]:
                if personId is None:
                    personId = str(personId).replace(" ", "")
                    colList.append("")
                elif infoMap.get(personId) is not None:
                    for info in infoMap.get(personId):
                        if info[0] == '用户手机号归属地':
                            colList.append(info[1])
                else:
                    colList.append("")
            df.loc[:, '用户手机号归属地'] = colList

    writer = pd.ExcelWriter(folderPath + "\\" + file)
    for key in data.keys():
        data[key].to_excel(writer, sheet_name=key, index=False)
    writer.save()

print("处理完成,回车键关闭!")
input()
