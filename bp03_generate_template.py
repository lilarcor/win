#!/usr/bin/env python
# -*- coding: utf-8 -*-
# @Author: Joey Lu
# @Date:   2017/6/12 15:22
# @Version: 1.1
import win32com.client
import copy
import time
import re
# import ctypes  # An included library with Python install.
import tkinter
import os

sep = os.sep  # 系统路径分隔符
dbpath = sep + sep + "shi01conofc01" + sep + "dept" + sep + "filing backup" + sep + "TA" + sep  # database path
scriptpath = dbpath + "sap_scripts" + sep  # script template path

dbfile = "SAP_TRAMMO_SH_BP.accdb"  # 数据库文件
tpfile = "bp03_generate_template"  # vbs模板文件文件名，不包含扩展名
tplist = []  # 将目标脚本文件名保存在列表
TIME = time.strftime("%Y-%m-%d-%H-%M", time.localtime())  # 用于生成文件名
keyword1 = "ID"  # 文件名关键字1,ID号
keyword2 = "search_term_ZH"  # 文件名关键字2,公司搜索名
keyword3 = "PIC"  # 文件名关键字3,提交人

record = dict()  # 生成字典记录数据库
LIST = []  # 保存字典值信息
KEY = []  # 保存字典的key值
head = []  # 保存数据库表头

conn = win32com.client.gencache.EnsureDispatch('ADODB.Connection')
DSN = 'PROVIDER = Microsoft.ACE.OLEDB.12.0;DATA SOURCE =' + dbpath + dbfile

conn.Open(DSN)
rs = win32com.client.Dispatch(r'ADODB.Recordset')
# 注意如果第一个记录为空，则 rs.MoveFirst() 会产生一个错误，如果此前将rs的 CursorLocation 设置为3，则此问题可解决
rs.CursorLocation = 3
rs_name = 'q_inactive'
# 不允许更新，用于查询
rs.Open('[' + rs_name + ']', conn, 1, 1)
fieldsQty = rs.Fields.Count  # 数据库表字段
recordQty = rs.RecordCount  # 数据库记录数
record_value_none = ("none", "na", "无")

# print(fieldsQty)
# print(recordQty)
rs.MoveFirst()
while True:
    if rs.EOF:
        break
    else:
        for i in range(0, fieldsQty):
            #   print(rs.Fields.Item(i).Name)
            #  print(rs.Fields.Item(i).Value)

            record = copy.deepcopy(record)
            # record[rs.Fields.Item(i).Name] = rs.Fields.Item(i).Value
            if str(rs.Fields.Item(i).Value).lower().strip() in record_value_none:
                record[rs.Fields.Item(i).Name] = ''
            elif rs.Fields.Item(i).Value:
                record[rs.Fields.Item(i).Name] = rs.Fields.Item(i).Value
            else:
                record[rs.Fields.Item(i).Name] = ''

            KEY.append(rs.Fields.Item(i).Name)

    LIST.append(record)
    rs.MoveNext()

conn.Close()

for i in range(0, recordQty):
    with open(scriptpath + tpfile + ".vbs", "r") as f1:
        s = f1.read()
    tplist.append(scriptpath + tpfile + "/" + TIME + "-" + str(LIST[i][keyword3]) + "-" + "ID" + str(LIST[i][keyword1]) + "-" + str(LIST[i][keyword2]) + ".vbs", )
    with open(tplist[i], 'w') as f2:
        # with open(scriptpath + tpfile + "/" + TIME + "-" + str(LIST[i][keyword3]) + "-" + "ID" + str(LIST[i][keyword1]) + "-" + str(LIST[i][keyword2]) + ".vbs", 'w') as f2:
        # print(f2.name)
        for v in LIST[i].keys():
            s = re.sub('\"' + str(v) + '\"', '\"' + str(LIST[i][v]) + '\"', s)  # 替换时搜索片段加入引号，避免变量与原文件里的内容重叠，造成错误覆盖
        f2.write(s)
    f2.close()
    f1.close()

clip = tkinter.Tk()
clip.clipboard_clear()
clip.clipboard_append(scriptpath + tpfile)  # 设置系统剪贴板内容
# ctypes.windll.user32.MessageBoxW(0, str(len(tplist)) + " records generated\nAnd path is in clipboard", "result", 0)
os.startfile(scriptpath + tpfile)
