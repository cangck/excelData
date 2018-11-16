#!coding:utf-8
# Author:pymingming

# -*- coding: utf-8 -*-
import os
import re
import uuid

import xlrd
import xlwt


def read_excel():
    email = []
    tel = []
    error = []
    emailcount = 0
    telcount = 0
    talcount = 0
    errorcount = 0

    dir = ['data1', 'data2', 'data3', 'data4']
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
                talcount = talcount + len(cols)
                for i in range(len(cols)):
                    # print("%s" % cols[i])
                    if re.match(
                            r'^[0-9a-zA-Z_.\-\+]{0,30}@[0-9a-zA-Z.\-\+]{1,20}\.[live,wohnmobil,novoa,sextl,info,com,cn,net,org,tv,edu,mx,io,tw,us,Com,gmx,at,smartsurv,co,za,inbox,lv,englpa,sextl,info,inbox,f-m,fm,biz,groeschel,wohnmobil,church]{1,3}$',
                            cols[i]):
                        email.append(cols[i])
                    elif re.match(r'^[0-9]*\-[0-9]*', cols[i]):
                        tel.append(cols[i])
                    else:
                        if re.match(r'^.*@.*$', cols[i]):
                            email.append(cols[i])
                        else:
                            error.append(cols[i])
                            # print(cols[i])

        if len(email) > 0:
            writeDate(email, 'email')
            emailcount = emailcount + len(email)
            email.clear()
        if len(tel) > 0:
            writeDate(tel, 'telephone')
            telcount = telcount + len(tel)
            tel.clear()

    if len(error) > 0:
        writeDate(error, 'error')

    print('email: ' + str(emailcount))
    print('tel: ' + str(telcount))
    print('error: ' + str(len(error)))

    print('counts: ' + str(talcount))
    print('totals: ' + str(emailcount + telcount + len(error)))


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
