#!user/bin/env python
# coding=utf-8

import xlrd


def readExcelDataByName(filename, sheetName):
    '''读取Excel文件和表名'''
    sheet = None
    errorMsg = None
    try:
        data = xlrd.open_workbook(filename)
        # sheet=data.sheet_by_index(0)#通过索引顺序获取，0表示第一张表
        # sheets = data.sheet_names()#获取文件中的表名
        sheet = data.sheet_by_name(sheetName)
        '''读取整张表并打印出来'''
        # for i in range(0,sheet.nrows):
        # 	row=sheet.row(i)
        # 	for j in range(0,sheet.ncols):
        # 		print(sheet.cell_value(i,j),"\t", end="")
        # 	print()
        # '''获取第几行的数据'''
        # print(sheet.row_values(0))
        # '''获取第n列的数据'''
        # print(sheet.col_values(1))
        i = j = k = m = n = 0
        # for v in sheet.col_values(0):
        #     if v == '正常类':
        #         i = i + 1
        #     elif v == "异常类":
        #         j = j + 1
        #     elif v == "业务规则":
        #         k = k + 1
        #     elif v == "主流程":
        #     	m=m+1
        #     elif v=="异常流":
        #     	n=n+1
        # print(filename + ' 中 ' + sheetName +
        #       '共有{}个正常类,{}个异常类,{}个业务规则,{}个主流程,{}个异常流'.format(i, j, k,m,n))
        for v in sheet.col_values(7):
        	if v=="通过":
        		i=i+1
        	elif v=="不通过":
        		j=j+1
        	elif v=="未执行":
        		k=k+1
        	elif v=="不执行":
        		m=m+1
        print(filename + ' 中 ' + sheetName +
            '测试用例通过了{}个,未通过的有{}个,未执行的有{}个,不执行的有{}个'.format(i, j, k,m))
    except Exception as msg:
        errorMsg = msg
    return sheet, errorMsg
if __name__ == '__main__':
    readExcelDataByName('检查点.xlsx', '测试用例初版-leangoo')
