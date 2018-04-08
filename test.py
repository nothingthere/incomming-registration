#!/usr/bin/python3
# test.py
# Author: Claudio <3261958605@qq.com>
# Created: 2018-04-07 22:31:46
# Code:
'''

'''
import openpyxl as xls

book = xls.load_workbook('x.xlsx')
sheets = book.sheetnames
sheet = book[sheets[0]]
print(sheet.title)
print('!!!')
