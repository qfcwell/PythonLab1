# -*- coding: utf-8 -*- 运行python
import pymssql
import os

def main():
    pass

def create_conn(count=10):
    with pymssql.connect(host='10.1.42.131', user="sa", password="qfc23834358Q",database="Product_Test") as conn:
        cur=conn.cursor()

def parse(path=''):
    pass

class P_Class():#产品类，族
    def __init__(self):
        self.Id=''
        self.Name=''
        self.Type=''
        self.Pic=''
        self.Table=''
        self.NoFill=[]
        self.MustFill=[]
        self.TypePara=[]
        self.EntPara=[]
        self.ParaList=[]

class P_Para():
    def __init__(self,display_name='',field_name='',metric='',unit=''):
        self.Id=''
        self.DisplayName=display_name
        self.FieldName=field_name
        self.Metric=metric
        self.Unit=unit
        self.Value=(display_name,field_name,metric,unit)
        self.RFieldName='##'.join([display_name,metric,unit])

    def set(self):
        with pymssql.connect(host='10.1.42.131', user="sa", password="qfc23834358Q",database="Product_Test") as conn:
            cur=conn.cursor()
            if self.Id:
                cur.execute(u"UPDATE P_Para SET DisplayName=")



class P_Category():
    def __init__(self,name=''):
        self.Id=''
        self.Name=name
        self.SubClass=[]
        self.Classes=[]

if __name__=="__main__":
    pass
