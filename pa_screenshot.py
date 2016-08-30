# -*- coding:utf-8 -*-
import sys, os, requests, time
import pandas as pd
import webbrowser as wb
import pyscreenshot
from PIL import Image

# 获取当前路径
cwd = os.getcwd()

link = []

xls = pd.ExcelFile('3.xlsx')
df = pd.read_excel(xls, header = None, index_col = None, na_value = None)

rows_df = len(df.index.values.tolist())


for i in range(rows_df):
    i_1 = i + 1
    wb.open(df.ix[i_1, 5],new = 0)
    time.sleep(10)
    r = requests.head(df.ix[i_1, 5])
    if r.status_code == 200:
        img = pyscreenshot.grab()
    else:
        continue
    img2 = img.crop((0,240,2540,1440))
    img2.save('{}.png'.format(i_1))