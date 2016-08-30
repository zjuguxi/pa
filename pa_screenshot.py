# -*- coding:utf-8 -*-
import sys, os, requests, time
import pandas as pd
import webbrowser as wb
import pyscreenshot
from PIL import Image
from docx import Document
from docx.shared import Inches

# 获取当前路径
cwd = os.getcwd()

# 自动截图并裁剪

xls = pd.ExcelFile('3.xlsx')
df = pd.read_excel(xls, header = None, index_col = None, na_value = None)

rows_df = len(df.index.values.tolist())

# 创建 Word 文档
document = Document()
document.add_heading('Document Title', 0)

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

    # 写入 Word 文档
    headline = df.ix[i_1, 3]
    p = document.add_heading('{}'.format(headline), level = 1)
    document.add_picture('{}.png'.format(i_1))
    document.save('report.docx')


# 生成 PPT

print('''I have fought the good fight,
         I have finished the race, 
         I have kept the faith.
         
         - 2 Timothy 4:7


         I've done my job, ma'am.
         : -)''')