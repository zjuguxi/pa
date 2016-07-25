import os
import urllib.request
import pandas
cwd = os.getcwd()

for i in list:

    request = urllib.request.Request('http://www.baidu.com')
    response = urllib.request.urlopen(request)

    f = open('./temp/{}', 'wb+').format(i)
    f.write(response.read())
