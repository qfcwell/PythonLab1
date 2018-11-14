# -*- coding: utf-8 -*-
import pymssql,cx_Oracle
import os

def get_staff():
    with pymssql.connect(host='10.1.1.117', user="auditor", password="hygj!@#456",database="cip") as conn:
        cur=conn.cursor()
        cur.execute(u"""SELECT LoginName,JobNo,Username,CompanyName,DeptName,Specialty FROM [dbo].[v_auditor_staff]""")
        res=cur.fetchall()
    return res

def drop_and_create_staff(oracle):
    with cx_Oracle.connect(oracle) as conn:
        cur=conn.cursor()
        try:
            cur.execute(u"DROP TABLE CAPOL_CIP_STAFF")
        except cx_Oracle.DatabaseError:
            pass
        cur.execute(u"CREATE TABLE CAPOL_CIP_STAFF(LoginName varchar2(50),JobNo varchar2(16),Username varchar2(50),CompanyName varchar2(50),DeptName varchar2(50),Specialty varchar2(50))")
        cur.execute(u"INSERT INTO CAPOL_CIP_STAFF(LoginName,JobNo,Username) VALUES('同步时间','0000',to_char(sysdate, 'yyyy-mm-dd hh24:mi:ss'))")
        conn.commit()

def check_and_insert_staff(oracle,staff):
    with cx_Oracle.connect(oracle) as conn:
        cur=conn.cursor()
        for (LoginName,JobNo,Username,CompanyName,DeptName,Specialty) in staff:
            cur.execute(u"SELECT LoginName,JobNo,Username,CompanyName,DeptName FROM CAPOL_CIP_STAFF WHERE JobNo=:1",[JobNo])
            res=cur.fetchone()
            if res:
                if (LoginName,JobNo,Username,CompanyName,DeptName)!=res:
                    cur.execute(u"UPDATE CAPOL_CIP_STAFF SET LoginName=:1,Username=:2,CompanyName=:3,DeptName=:4 WHERE JobNo=:5",[LoginName,Username,CompanyName,DeptName,JobNo])
            else:
                cur.execute(u" INSERT INTO CAPOL_CIP_STAFF(LoginName,JobNo,Username,CompanyName,DeptName,Specialty) VALUES(:1,:2,:3,:4,:5,:6)",[LoginName,JobNo,Username,CompanyName,DeptName,Specialty])
        cur.execute(u"UPDATE CAPOL_CIP_STAFF SET Username=to_char(sysdate, 'yyyy-mm-dd hh24:mi:ss') where JobNo='0000'")
        conn.commit()


def run():
    oracle='steven/qfc23834358@10.1.42.110:1521/xe'
    #oracle='ccd/hygj1234@10.1.1.213:1521/orcl'
    #drop_and_create(oracle)
    staff=get_staff()
    #drop_and_create_staff(oracle)
    check_and_insert_staff(oracle,staff)

def test():
    with pymssql.connect(host='10.1.1.117', user="auditor", password="hygj!@#456",database="cip") as conn:
        cur=conn.cursor()
        cur.execute(u"""SELECT LoginName,JobNo,Username,CompanyName,DeptName,Specialty FROM [dbo].[v_auditor_staff] where LoginName=%s """,('wuzhuwei'))
        res=cur.fetchall()
    print(res)


if __name__=="__main__":
    #run()
    #test()
    pass
