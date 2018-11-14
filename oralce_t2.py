# -*- coding: utf-8 -*- 运行python
import pymssql,cx_Oracle
import os


os.environ['NLS_LANG'] = 'SIMPLIFIED CHINESE_CHINA.UTF8'
oracle='steven/qfc23834358@10.1.42.131:1521/xe'

def main():
    pass

def create_conn(count=10):
    lst=[]
    for i in range(count):
        conn=cx_Oracle.connect(oracle)
        lst.append(conn)
    return lst

if __name__=="__main__":
    t1=create_conn()

t1=create_conn()
