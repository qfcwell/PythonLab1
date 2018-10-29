# -*- coding: utf-8 -*-
import pymssql,cx_Oracle
import os

def get_project(host):
    with pymssql.connect(host=host, user="readonly", password="capol!@#456",database="CAPOL_Project") as conn:
        cur=conn.cursor()
        cur.execute(u"""SELECT PrjNo,PrjName,CreateTime,Project.FolderName+'_'+PrjPhase.FolderName AS FolderName ,PrjPhase.PrjPhaseName
            FROM CAPOL_Project.dbo.Project left join CAPOL_Project.dbo.PrjPhase on PrjPhase.PrjId = Project.Id 
            where Project.RecordState='A' and PrjPhase.RecordState='A'""")
        res=cur.fetchall()
    return res

def drop_and_create(oracle):
    with cx_Oracle.connect(oracle) as conn:
        cur=conn.cursor()
        cur.execute(u"DROP TABLE CAPOL_PRJACSSERVER")
        cur.execute(u"CREATE TABLE CAPOL_PRJACSSERVER(PRJ_NO varchar2(16),PRJ_NAME varchar2(1000),PRJ_PHASE_NAME varchar2(16),CREATE_TIME DATE,S_NAME varchar2(10),SAP_PRJNO varchar2(16),FOLDER varchar2(1000))")
        cur.execute(u"INSERT INTO CAPOL_PRJACSSERVER(PRJ_NO,PRJ_NAME,CREATE_TIME,S_NAME,SAP_PRJNO) VALUES('0000','最新同步时间',TRUNC(SYSDATE),'0000','0000')")
        conn.commit()

def check_and_insert(oracle,Projects,S_NAME):
    with cx_Oracle.connect(oracle) as conn:
        cur=conn.cursor()
        for (PrjNo,PrjName,CreateTime,FolderName,PrjPhaseName) in Projects:
            FolderName="D:\\AcsModule\\WorkFolder\\Project\\"+FolderName
            cur.execute(u"SELECT PRJ_NAME,PRJ_PHASE_NAME,CREATE_TIME,S_NAME,FOLDER FROM CAPOL_PRJACSSERVER WHERE PRJ_NO=:1 AND PRJ_PHASE_NAME=:2",[PrjNo,PrjPhaseName])
            res=cur.fetchone()
            if res:
                (PRJ_NAME,PRJ_PHASE_NAME,CREATE_TIME,S,FOLDER)=res
                if (PrjName,PrjPhaseName,CreateTime,S_NAME,FolderName)!=res and CreateTime >= CREATE_TIME:
                    cur.execute(u"UPDATE CAPOL_PRJACSSERVER SET PRJ_NAME=:1,PRJ_PHASE_NAME=:2,CREATE_TIME=:3,S_NAME=:4,FOLDER=:5 where PRJ_NO=:6",[PrjName,PrjPhaseName,CreateTime,S_NAME,FolderName,PrjNo])
            else:
                cur.execute(u"INSERT INTO CAPOL_PRJACSSERVER(PRJ_NO,PRJ_NAME,PRJ_PHASE_NAME,CREATE_TIME,S_NAME,SAP_PRJNO,FOLDER) VALUES(:1,:2,:3,:4,:5,:6,:7)",[PrjNo,PrjName,PrjPhaseName,CreateTime,S_NAME,PrjNo,FolderName])
        cur.execute(u"UPDATE CAPOL_PRJACSSERVER SET PRJ_NAME='最新同步时间',CREATE_TIME=TRUNC(SYSDATE),S_NAME='0000' where PRJ_NO='0000'")
        conn.commit()


def run():
    lst=[
    ('深圳','10.1.246.1'),
    ('广州','10.2.1.114'),
    ('长沙','10.3.1.3'),
    ('上海','10.6.1.240')
    ]
    #oracle='steven/qfc23834358@10.1.42.66:1521/xe'
    oracle='ccd/hygj1234@10.1.1.213:1521/orcl'
    #drop_and_create(oracle)
    for (S_NAME,host) in lst:
        Projects=get_project(host)
        check_and_insert(oracle,Projects,S_NAME)



if __name__=="__main__":
    run()
