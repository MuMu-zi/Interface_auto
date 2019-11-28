# !/usr/bin/python
# -*- coding: utf-8 -*-
# @File   : get_case.py
# @time   : 2019-11-25 20:14
# @Author : yangmuzi
'''
Excel表操作
'''

import xlrd     # 读
import xlwt     # 写

file_path = r'/Users/yanghuan/PycharmProjects/Interface_auto/test_example.xlsx'


#
# class OpCase(object):
#
#     def g_case(self,case_path):
cases = []
# 获取book对象
book = xlrd.open_workbook(file_path)

# 两种方式获取工作表
# sheet = book.sheet_by_index(0)
sheet = book.sheet_by_name('工作表1')

# sheets = book.sheets()    # 获取所有的工作表对象

rows = sheet.nrows      # 获取行数
cols = sheet.ncols      # 获取列数
# print(sheet.row_values(1))      # 获取指定行
# print(sheet.col_values(0))      # 获取指定列
# print(sheet.cell_value(0,0))    # 获取指定单元格的内容

# # 获取所有行,按行读取
# for row in range(rows):
#     print(sheet.row_values(row))

# # 获取所有列，按列读取
# for col in range(cols):
#     print(sheet.col_values(col))

# 将首行与其他行组成字典





for i in range(1,rows):
    for j in range(sheet.ncols):
        row_data = sheet.row_values(i,j)
        cases.append(row_data)
# print(cases)

        # return cases




# if __name__ == '__main__':
#     case_path = 'test_example.xlsx'
#
#







