#!/usr/bin/env Python
# coding=utf-8
import pandas as pd
import codecs
xd = pd.ExcelFile('D:\\1.xlsx')
df = xd.parse()
print(df.columns)
print(df.values)
#res=df.to_html(header = True,index = False)
#print(res)

def read_excel()