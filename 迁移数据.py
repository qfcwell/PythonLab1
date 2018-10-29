# -*- coding: utf-8 -*- 运行python
import pymssql
import pypyodbc

def get_access_data():
    str = 'Driver={Microsoft Access Driver (*.mdb,*.accdb)};DBQ=D:\\文档\\GitHub\\Python_old\\database2.accdb'
    with pypyodbc.win_connect_mdb(str) as conn:
        cur=conn.cursor()
        cur.execute(u"SELECT top 100 * FROM files")
        return cur.fetchall()

def migrate():
    pass

def run():
    for line in get_access_data():
        print(line)

if __name__=="__main__":
    run()