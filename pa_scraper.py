# -*- coding:utf-8 -*-

import requests
import sys
import pandas as pd


link = []

def download_web(x):
    ff = open('%s.html', ) % x
    ff.writelines(x.text)
    ff.close()

for i in link:
    r = requests.get(i)
    download_web(r)

