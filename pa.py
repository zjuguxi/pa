import os
import numpy as np
import pandas as pd
import xlsxwriter

# 得到当前path
cwd = os.getcwd()

# 打开 Excel 文档的 Pickup 汇总表，读取链接
xls = pd.ExcelFile('1.xls')
df = pd.read_excel(xls, 'Pickup', header = [43], index_col = None, na_value = None)

# 打开 Realeases 读取文章标题
df_headline = pd.read_excel(xls, 'Releases', header = None, index_col = 0, na_value = None)

# 把 Releases 的 columns 改成 A/B
df_headline.columns = ['A', 'B']

# Releases 的总行数
n = df_headline.index.values.tolist()[-1]
series_headline = []
series_storynumber = []
for i in n:
    if df_headline.ix[:, 'A'] == 'Headline':
        series_headline.append(df_headline.ix[:, 'B'])
    if df_headline.ix[:, 'A'] == 'Story Number':
        series_storynumber.append(df_headline.ix[:, 'B'])

df_releases = DataFrame(series_headline, index = series_storynumber)

'''

df_releases_column = ['Headline', 'Language', 'Realease Clear Time', 'Story Number', 'View Release on '] 



s_language_releases = df_headline.loc['Language', :]
s_topic_releases = df_headline.loc['Headline', :]


dict_headline_language = {'Language': s_language_releases,
                                         'Headline': s_topic_releases}
df_new_1 = pd.DataFrame([dict_headline_language])
# 提取row为 Headline 的column
print(df_new_1)

# print(df_headline.iloc['Headline', 1])

'''
url_prefix = 'http://'

for i in df:
    if url_prefix in i:
        print(i)
'''
# 打印所有行的第4列（从0开始计数）
# print(df.iloc[:, 4])

# 提取链接、语言、媒体

s_link = df.iloc[:, 4]
s_language = df.iloc[:, 1]
s_media = df.iloc[:, 2]

dict_link= {'Link': s_link,
                  'Language': s_language,
                  'Media': s_media}
print(dict_link)
writer = pd.ExcelWriter('2.xlsx', engine='xlsxwriter')
df_new = pd.DataFrame(dict_link)
df_new.to_excel(writer, sheet_name='Sheet1')
writer.save()


df_2 = pd.read_excel('2.xlsx', header = None, index_col = None, na_value = None)
#print(df_2.loc[:, 1])

# 每个链接生成字典，包含『Headline』和『View Release on』两项，然后加入列表 links[]

writer = pd.ExcelWriter('3.xlsx', engine='xlsxwriter')
df_new_1.to_excel(writer, sheet_name='Sheet1')
writer.save()



for i in list:

    request = urllib.request.Request('http://www.baidu.com')
    response = urllib.request.urlopen(request)

    f = open('./temp/{}', 'wb+').format(i)
    f.write(response.read())
'''