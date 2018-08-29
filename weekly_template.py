# -*- coding: utf-8 -*-
from collections import OrderedDict
from pyexcel_xlsx import save_data
import xlrd
import pymysql.cursors
import os
DB_HOST = "***.**.**.*8"
DB_PORT =222
DB_USER = "222"
DB_PASSWORD = "222"
DB_NAME = "222"

# Main routine
# 写Excel数据, xls格式,以sheet为单位
def save_xls_file(xls_header,xls_sheet, file_name):
    sheet_data = []
    data = OrderedDict()
    if os.path.exists(file_name):
        #d打开表格文件
        with xlrd.open_workbook(file_name) as rb:
          #获取表格
          sheet = rb.sheet_by_index(0)
          #行数
          nrows = sheet.nrows
          #遍历将每一行内容加入sheet_data
          for i in range(nrows):
            sheet_data.append(sheet.row_values(i))
        #每次再次写入内容之前先插入一个空白行
        list_space = []
        sheet_data.append(list_space)
     #添加表头
    sheet_data.append(xls_header)
    for row_data in xls_sheet:
         sheet_data.append(row_data)
    # 添加sheet表
    data.update({u"Sheet1": sheet_data})
    print(type(data))
    # 保存成xls文件
    save_data(file_name,data)

# 读取txt文本  表头以及sql
f = open(r'F:\LWtest\header.txt','r',encoding='utf-8')
lines = f.readlines()
#names = locals()
sheetLists=[]
for line in lines:
    sheetLists.append(list(map(str,line.split('\n'))))
    print(sheetLists)

f = open(r'F:\LWtest\sql.txt','r',encoding='utf-8')
lines=f.readlines()
sqlLists=[]
for line in lines:
    #用每一行末尾的换行符做分隔 保证sql语句的完整 可以替换任何不会使用的符号替代
    sqlLists.append(list(map(str,line.split('\n'))))

#数据库连接
db = pymysql.connect(host=DB_HOST,
                     user=DB_USER,
                     port=DB_PORT,
                     passwd=DB_PASSWORD,
                     # init_command="set names utf8",
                     cursorclass=pymysql.cursors.SSDictCursor,
                     charset='utf8')
db.select_db(DB_NAME)
cur = db.cursor(cursor=pymysql.cursors.SSCursor)
#遍历表头list用来确定io操作的次数
try:
  for i in range(len(sheetLists)):
      sheet=sheetLists[i][0].split()#将每一行的表头以空格做分隔为单个字段
      rowlist=[]
      sql=sqlLists[i][0]
      print(sql)
      cur.execute(sql)
      rows=cur.fetchall()
      save_xls_file(sheet,rows,r"F:\LWtest\result3.xlsx")
except:
    print("异常")
cur.close()
print("程序执行完毕")


