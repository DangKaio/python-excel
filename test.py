#!user/bin/env python
# coding=utf-8
# @Author  : Dang
# @Time    : 2018/5/22 17:25
# @Email   : 1370465454@qq.com
# @File    : Test_SplitExcel.py
# @Description:对excel单元格已；进行分隔，完成后需要对
import xlrd
import xlwt
import os
import re


def readExcelDataByName(filename, sheetName, num,save_filename):
    """
        :param filename:输入文件路径和名字+后缀
        :param sheetName:输入表名
        :param num:输入要分隔的列
        :param save_filename:要保存的文件名称
    """
   
    wb = xlrd.open_workbook(filename)
    # sheet=data.sheet_by_index(0)#通过索引顺序获取，0表示第一张表
    # sheets = data.sheet_names()#获取文件中的表名
    sheet = wb.sheet_by_name(sheetName)
    ncols = sheet.ncols
    # print(ncols)
    # 获取行数
    nrows = sheet.nrows
    print("nrows %d, ncols %d" % (nrows, ncols))
    row_list = []
    work_book = xlwt.Workbook()
    sheet1 = work_book.add_sheet(sheetName[:-1], cell_overwrite_ok=True)
    # for m in range(0, nrows):
    #     for n in range(0, num):  # 列
    #         data_init = sheet.cell_value(m, n)
    #         sheet1.write(m, n, data_init)
    k=0
    for m in range(1, nrows):
        data = sheet.cell_value(m, num)
        row_list = re.split("；", data.replace("\n", ""))
        for n in range(0, len(row_list)):  # 列
            sheet1.write(n + num, m, row_list[n])  # 从第6列写入
        k=k+len(row_list)
        # print(m)
        # print(row_list)
        row_list.clear()
    # print(k-nrows)

    print("大约有 %d个用例，此处只做大概统计，具体需要根据实际情况减去相应值" %(k-nrows))

    for i in range(num, ncols):
        data_add = sheet.cell_value(0, i)
        # print(kapian)
        sheet1.write(0, i, data_add)
    if os.path.exists(save_filename + sheetName + ".xls"):
        os.remove(save_filename + sheetName + ".xls")
        work_book.save(save_filename + sheetName + ".xls")
    else:
        work_book.save(save_filename + sheetName + ".xls")
    print("转换完成，请查看 %s%s.xls的文档。" %(save_filename,sheetName))
if __name__ == '__main__':
    """
        :param filename:输入文件路径和名字+后缀
        :param sheetName:输入表名
        :param num:输入要分隔的列
        :param save_filename:要保存的文件路径和名称，要保存文件名会和表名自动组合形成新的文件 如：分隔结果文档Sheet1.xls
    """
    readExcelDataByName('F:\\项目\\华润三九\\华润三九测试用例-6.7.xlsx', '医生端小程序初版', 5,"F:\\项目\\华润三九\\分隔结果文档")
