#!coding:utf-8
# Author:pymingming

# -*- coding: utf-8 -*-
import os
import re
import uuid

import xlrd
import xlwt
import json
import random


def read_excel():
    data = []
    talcount = 0
    name = 0
    # dir = ['data1', 'data2', 'data3', 'data4']
    dir = ['data1']
    for h in range(len(dir)):
        print(dir[h])
        for dirpath, dirnames, filenames in os.walk(dir[h]):
            for filepath in filenames:
                print(os.path.join(dirpath, filepath))

                # 打开文件
                workbook = xlrd.open_workbook(os.path.join(dirpath, filepath))
                # 获取所有sheet
                # print(workbook.sheet_names())  # [u'sheet1', u'sheet2']
                # 获取sheet1
                sheet2_name = workbook.sheet_names()[0]
                # print(sheet2_name)
                # 根据sheet索引或者名称获取sheet内容
                sheet2 = workbook.sheet_by_name(sheet2_name)
                # sheet的名称，行数，列数
                cols = sheet2.col_values(9)  # 获取列内容
                for i in range(len(cols)):
                    if i > 20:
                        break
                    if re.match(r'^[0-9]*\-[0-9]*', cols[i]):
                        tel = {}
                        res = re.search('(.*)-(.*)', cols[i])
                        tel['name'] = name
                        tel['tel'] = "+" + res.group(1) + res.group(2)
                        data.append(tel)
                        name = name + 1

        print(json.dumps(tel))
        fp = open("test.txt", 'w')
        fp.write(json.dumps(data));
        fp.close()


def writeDate(file, name):
    wbk = xlwt.Workbook()
    sheet = wbk.add_sheet('sheet')

    k = 0
    data = list(set(file))
    lenth = len(data)
    for i in range(lenth):
        sheet.write(int(i / 20), i % 20, data[i])  # 循环写入每行数据

    wbk.save("target/" + name + str(uuid.uuid4()) + '.xls')  # 保存excel必须使用后缀名是.xls的，不是能是.xlsx的


if __name__ == '__main__':
    read_excel()
