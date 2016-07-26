import os
import urllib.request
import pandas as pd

# 链接池目前为空
links = []

# 得到当前path
cwd = os.getcwd()

# 打开 Excel 文档，读取链接
xls = pd.ExcelFile('1.xls')
df = pd.read_excel(xls, 'Releases', header = [1], index_col = ['headline'], na_value = None)





print(df)

'''
for i in list:

    request = urllib.request.Request('http://www.baidu.com')
    response = urllib.request.urlopen(request)

    f = open('./temp/{}', 'wb+').format(i)
    f.write(response.read())
'''