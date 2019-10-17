# -*- coding:utf-8-*-
import excelXarray

# Excel to Array
array = excelXarray.excelToArray('test.xlsx')
print(array)

# Array to Excel
array = [['A', 'B', 'C'], [1.0, 'En', '中文'], [2.0, 'de', 'ru'], [3.0, 'Fr', 'ja'], ['', 'h', '']]
excelXarray.arrayToExcel(array, 'export.xlsx')
