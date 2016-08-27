import os
import numpy as np
import pandas as pd
import xlsxwriter

# 得到当前path
cwd = os.getcwd()

# 打开 Excel 文档的 Pickup 汇总表，读取链接
xls = pd.ExcelFile('1.xls')

# df = pd.read_excel(xls, 'Pickup', header = 43, index_col = None, na_value = None)

df_pickup = pd.read_excel(xls, 'Pickup', header = None, index_col = None, na_value = None)

# Pickup 的总 row 数
p = len(df_pickup.index.values.tolist())

# 找到 Pickup column[0] 里 Story Number 第一次出现的位置，取其 index，命名为 m
# 从 Story Number 处开始截取 DataFrame
s = df_pickup.iloc[:, 0]
m = s[s == 'Story Number'].index[0]
df = pd.read_excel(xls, 'Pickup', header = m, index_col = None, na_value = None)

# 找到 Twitter 的 index，命名为 twt
# 从 Twitter Handle 处开始截取 DataFrame
twt = s[s == 'Twitter Handle'].index[0]
df_twitter = pd.read_excel(xls, 'Pickup', header = twt, index_col = None, na_value = None)
# 把 Twitter 信息生成新表格
writer = pd.ExcelWriter('3_twitter.xlsx', engine='xlsxwriter')
df_twitter.to_excel(writer, sheet_name='Sheet1')
writer.save()

# 打开 Realeases 读取文章标题
df_headline = pd.read_excel(xls, 'Releases', header = None, index_col = 0, na_value = None)

# Pickup 的总 row 数
p = len(df.index.values.tolist())
# Releases 的总 row 数
n = len(df_headline.index.values.tolist())

# 提取 Releases 里对应的 Story Number和 Headline，形成 Series
series_headline = []
series_storynumber = []

for i in range(n):
    if df_headline.index[i] == 'Headline':
        series_headline.append(df_headline.ix[i, 1])
    if df_headline.index[i] == 'Story Number':
        series_storynumber.append(df_headline.ix[i, 1])

df_releases = pd.DataFrame(series_headline, index = series_storynumber) # 将2个series合并为一个dataframe

# Releases新表格 的总 row 数
q = len(df_releases.index.values.tolist())

print(df_releases)


list_addon = []

for i in range(p):
    for j in range(q):
        if df.ix[i, 'Story Number'] == df_releases.index[j]: # 对比新 Releases 和 Pickup 的Story Number
            df.ix[i, 'Headline'] = df_releases.ix[j, 0]


writer = pd.ExcelWriter('3.xlsx', engine='xlsxwriter')
df.to_excel(writer, sheet_name='Sheet1')
writer.save()