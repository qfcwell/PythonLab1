# -*- coding: utf-8 -*- 运行python
import pymssql,cx_Oracle

def run():
    oracle='ccd/hygj1234@10.1.1.213:1521/orcl'
    with cx_Oracle.connect(oracle) as conn:
        cur=conn.cursor()
        cur.execute(u"select 工程编号 from Capol_ProjectOpinions_2018")
        res=cur.fetchone()
        print(res)



if __name__=="__main__":
    run()


