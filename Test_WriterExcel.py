#!user/bin/env python
# coding=utf-8

import xlrd
import xlwt
import os
import time
import sys
sys.path.append('../')
from openpyxl import load_workbook

strtime = time.strftime('%Y-%m-%d_%H_%M_%S')


def readExcelDataByName(filename, sheetName):
    '''读取Excel文件和表名'''
    wb = xlrd.open_workbook(filename)
    # sheet=data.sheet_by_index(0)#通过索引顺序获取，0表示第一张表
    # sheets = data.sheet_names()#获取文件中的表名
    sheet = wb.sheet_by_name(sheetName)
    ncols = sheet.ncols
    # 获取行数
    nrows = sheet.nrows
    print("nrows %d, ncols %d" % (nrows, ncols))
    row_list = []
    work_book = xlwt.Workbook()
    sheet1 = work_book.add_sheet(sheetName[:-9])
    k = 1
    for j in range(1, nrows):
        # 获取单元格
        for i in range(6, ncols):
            data = sheet.cell_value(j, i)
            if data == "":
                continue
            else:
                row_list.append(str(k) + "." + data + "\n")
            k = k + 1
        # print(row_list)
        sheet1.write(j, 6, row_list)
        row_list.clear()
        k = 1
    for m in range(0, nrows):
        for n in range(0, 6):  # 列
            data = sheet.cell_value(m, n)
            sheet1.write(m, n, data)
    if os.path.exists("结果" + sheetName + ".xls"):
        os.remove("结果" + sheetName + ".xls")
        work_book.save("结果" + sheetName + ".xls")
    else:
        work_book.save("结果" + sheetName + ".xls")
if __name__ == '__main__':
    readExcelDataByName('F:\\项目\\华润三九\\华润三九测试用例-6.7.xlsx', '代表小程序初版--leangoo')
    readExcelDataByName('F:\\项目\\华润三九\\华润三九测试用例-6.7.xlsx', '医生端小程序-leangoo')
