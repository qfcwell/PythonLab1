# -*- coding: utf-8 -*-
import pymssql,sqlite3,pypyodbc

def main():
    with pymssql.connect(host="10.1.42.102", user="qifangchen", password="qfc23834358",database="test2") as conn:
        cur=conn.cursor()
        cur.execute(u"select top 1000000 id,file_path_name from files4")
        with sqlite3.connect('数据源.sqlite3') as cur2:
            cur2.execute(u"CREATE TABLE files(fid INT PRIMARY KEY NOT NULL,file_path_name TEXT)")
            while 1:
                res=cur.fetchone()
                if res:
                    (fid,file_path_name)=res
                    cur2.execute(u"insert into files(fid,file_path_name) values(?,?)",(fid,file_path_name))
                else:
                    break
def main2():
    with pymssql.connect(host="10.1.42.102", user="qifangchen", password="qfc23834358",database="test2") as conn:
        cur=conn.cursor()
        cur.execute(u"select top 1000000 id,file_path_name from files4")
        str = 'Driver={Microsoft Access Driver (*.mdb)};DBQ=D:\\GitHub\\lab\\Python\\数据源.mdb'
        with pypyodbc.win_connect_mdb(str) as conn2:
            cur2 = conn2.cursor()
            cur2.execute(u"CREATE TABLE files(fid INT PRIMARY KEY NOT NULL,file_path_name TEXT)")
            while 1:
                res=cur.fetchone()
                if res:
                    (fid,file_path_name)=res
                    cur2.execute(u"insert into files(fid,file_path_name) values(?,?)",(fid,file_path_name))
                else:
                    break

if __name__ == '__main__':
    main2()
  