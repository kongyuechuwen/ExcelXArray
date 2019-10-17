# -*- coding:utf-8-*-
'''
@zhengwen 610756827@qq.com
@2019年10月12日
excel:
A B C
1 a a
2 b b
3 c c
Array[][]
[['A','B','C'],['1','a','a'],['2','b','b'],['3','c','c']]
读取excel表格，返回一个二维数组
'''

import xlrd
import xlsxwriter

# default excel sheetname is 'Sheet1'
def excelToArray(path, *sheetname):
    book = xlrd.open_workbook(path)
    sn = 'Sheet1'
    if(len(sheetname) > 0):
        sn = sheetname[0]
    sheet = book.sheet_by_name(sn)
    data = [[sheet.cell_value(r, c) for c in range(sheet.ncols)] for r in range(sheet.nrows)]
    return data

# array to excel
'''
Array[][]
[['A','B','C'],['1','a','a'],['2','b','b'],['3','c','c']]
excel:
A B C
1 a a
2 b b
3 c c
'''
def arrayToExcel(array, path):
    workbook = xlsxwriter.Workbook(path)
    worksheet = workbook.add_worksheet()
    row = 1
    # print(array)
    for data in array:
        worksheet.write_row('A'+str(row),data)
        row +=1
    workbook.close()
