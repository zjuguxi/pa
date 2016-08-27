# -*- coding:utf-8 -*-

import sys, os
import pandas as pd

xls = pd.ExcelFile('3.xlsx')
df = pd.read_excel(xls, header = None, index_col = None, na_value = None)

# 获取当前路径
cwd = os.getcwd()

link = []

