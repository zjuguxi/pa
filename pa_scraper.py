# -*- coding:utf-8 -*-
import sys, os, requests
import pandas as pd
import webbrowser as wb
import pyscreenshot

# 获取当前路径
cwd = os.getcwd()

link = []

xls = pd.ExcelFile('3.xlsx')
df = pd.read_excel(xls, header = None, index_col = None, na_value = None)

rows_df = len(df.index.values.tolist())

for i in range(rows_df):
    wb.open(df.ix[i + 1, 5])
    r = requests.head(df.ix[i, 5])
    if r.status_code == 200
        pyscreenshot.grab_to_file(/screenshot/%s.png) % i + 1
    else:
        break








