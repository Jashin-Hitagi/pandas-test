import pymysql
import pandas as pd

user = input("请输入数据库账号: ")
pwd = input("请输入数据库密码: ")
dataBase = input("请输入数据库名: ")
print("输入一级搜索条件(搜索字段名，条件字段名，表名，条件)，英文逗号分隔: ")
keyWord = input().split(',')
print("输入二级搜索字段名: ")
header = input()
print("输入二级搜索表名，英文逗号分隔: ")
tableList = input().split(',')
print("输入结果导出路径: ")
filePath = input() + "\\" + "result.xlsx"

keySql = '''
select %s from %s where %s = '%s'
''' % (keyWord[0], keyWord[2], keyWord[1], keyWord[3])

# 打开数据库连接
db = pymysql.connect(host="localhost", user=user, password=pwd, database=dataBase)
# db = pymysql.connect(host="localhost", user='root', password='123456', database='world')

# 使用cursor()方法创建一个游标对象cursor
cursor = db.cursor()

# 使用execute()方法执行SQL查询
cursor.execute(keySql)
keywordList = []
data = cursor.fetchall()
for word in data:
    keywordList.append(word[0])
keyword_str = '(' + ",".join('\'' + str(i) + '\'' for i in keywordList) + ')'

writer = pd.ExcelWriter(filePath)
for table in tableList:
    resultSql = '''
    select * from %s where %s in %s
    '''%(table,header,keyword_str)
    cursor.execute(resultSql)
    data = cursor.fetchall()
    des = cursor.description
    # 将数据truple转换为DataFrame
    col = []
    for i in des:
        col.append(i[0])
    data = list(map(list, data))
    data = pd.DataFrame(data, columns=col)
    data.to_excel(writer, sheet_name=table, index=False)
writer.save()

# 关闭数据库连接
db.close()
print("处理完成,回车键关闭!")
input()
