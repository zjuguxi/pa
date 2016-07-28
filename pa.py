import os
import urllib.request
import pandas as pd
import numpy as np

# 链接池dict目前为空
link1 = {}
link1['headline1'] = 'South China sea'
link1['View Realease on'] = 'http://'
links = [link1]

# 得到当前path
cwd = os.getcwd()

# 打开 Excel 文档，读取链接
xls = pd.ExcelFile('1.xls')
pickup = pd.read_excel(xls, 'Pickup', header = [44], index_col = None, na_value = None)

df = pickup.DataFrame(column = list('ABCDEFG'))

# 每个链接生成字典，包含『Headline』和『View Release on』两项，然后加入列表 links[]



print(df)

# 循环爬取链接并保存本地
'''
for i in list:

    request = urllib.request.Request('http://www.baidu.com')
    response = urllib.request.urlopen(request)

    f = open('./temp/{}', 'wb+').format(i)
    f.write(response.read())
'''

# 循环截图