# -*- coding: utf-8 -*-
import pymssql,cx_Oracle
import os


os.environ['NLS_LANG'] = 'SIMPLIFIED CHINESE_CHINA.UTF8'
acs_server_lst={'深圳':'10.1.246.1','广州':'10.2.1.114','长沙':'10.3.1.3','上海':'10.6.1.15','海南':'10.14.2.10'}

def run():
    method='test'
    #method='capol'
    drop=0
    with acs_project_sync(method=method) as s:
        s.run_sync(drop=drop)
    with staff_sync(method=method) as s:
        s.run_sync(drop=drop)
    with project_role_sync(method=method) as s:
        s.run_sync()

class sync():
    def __init__(self,method='test'): 
        self.cip=pymssql.connect(host='10.1.1.117', user="cip_ro", password="hygj!@#4567",database="cip")
        self.oracle_ccd=cx_Oracle.connect('ccd/hygj1234@10.1.1.213:1521/orcl')
        self.oracle_test=cx_Oracle.connect('steven/qfc23834358@10.1.42.131:1521/xe')
        if method=='test':
            self.oracle=self.oracle_test
        else:
            self.oracle=self.oracle_ccd
        self.acs_server={}
        self.acs_server_lst=acs_server_lst
        for S_NAME in acs_server_lst:
            self.acs_server[S_NAME]=pymssql.connect(host=acs_server_lst[S_NAME], user="readonly", password="capol!@#456",database="CAPOL_Project")

    def __enter__(self):
        return self

    def __exit__(self, type, value, trace):
        self.close()

    def close(self):
        self.a=1
        self.cip.close()
        self.oracle_ccd.close()
        self.oracle_test.close()
        for S_NAME in self.acs_server:
            self.acs_server[S_NAME].close()


class acs_project_sync(sync):
    def __init__(self, *arg, **kw):
        sync.__init__(self, *arg, **kw)
        self.Projects=[]
    
    def get_comp_dept(self,PrjNo):
        cur=self.cip.cursor()
        cur.execute('''SELECT org.fd_name as comp,dept.fd_name as dept from wt_project_info prj
            left join sys_org_element org on org.fd_id=prj.doc_company_id 
            left join sys_org_element dept on dept.fd_id=prj.doc_dept_id
            WHERE prj.fd_project_no=%s''',(PrjNo))
        return cur.fetchone()

    def get_project(self):
        for S_NAME in self.acs_server:
            cur=self.acs_server[S_NAME].cursor()
            cur.execute(u"""SELECT PrjNo,PrjName,'','',CreateTime,Project.FolderName+'_'+PrjPhase.FolderName AS FolderName ,
                PrjPhase.PrjPhaseName,PrjPhase.id as PrjPhaseId
                FROM CAPOL_Project.dbo.Project left join CAPOL_Project.dbo.PrjPhase on PrjPhase.PrjId = Project.Id 
                where Project.RecordState='A' and PrjPhase.RecordState='A'""")
            res=cur.fetchall()
            for (PrjNo,PrjName,comp,dept,CreateTime,FolderName,PrjPhaseName,PrjPhaseId) in res:
                comp_dept=self.get_comp_dept(PrjNo)
                if comp_dept:
                    (comp,dept)=comp_dept
                self.Projects.append((PrjNo,PrjName,comp,dept,CreateTime,FolderName,PrjPhaseName,PrjPhaseId,S_NAME))
        return self.Projects

    def drop_and_create(self):
        cur=self.oracle.cursor()
        try:
            cur.execute(u"DROP TABLE CAPOL_PRJACSSERVER")
        except cx_Oracle.DatabaseError:
            pass
        cur.execute(u"CREATE TABLE CAPOL_PRJACSSERVER(PRJ_NO varchar2(16),PRJ_NAME varchar2(1000),PRJ_COMP varchar2(48),PRJ_DEPT varchar2(48),PRJ_PHASE_NAME varchar2(16),PrjPhaseId varchar2(16),CREATE_TIME DATE,S_NAME varchar2(10),SAP_PRJNO varchar2(16),FOLDER varchar2(1000))")
        cur.execute(u"INSERT INTO CAPOL_PRJACSSERVER(PRJ_NO,PRJ_NAME,CREATE_TIME,S_NAME,SAP_PRJNO) VALUES('0000','最新同步时间',TRUNC(SYSDATE),'0000','0000')")
        self.oracle.commit()

    def check_and_insert(self):
        cur=self.oracle.cursor()
        for (PrjNo,PrjName,comp,dept,CreateTime,FolderName,PrjPhaseName,PrjPhaseId,S_NAME) in self.Projects:
            FolderName="D:\\AcsModule\\WorkFolder\\Project\\"+FolderName
            cur.execute(u"SELECT PRJ_NAME,PRJ_COMP,PRJ_DEPT,PRJ_PHASE_NAME,CREATE_TIME,S_NAME,FOLDER FROM CAPOL_PRJACSSERVER WHERE PRJ_NO=:1 AND PrjPhaseId=:2",[PrjNo,PrjPhaseId])
            res=cur.fetchone()
            if res:
                (PRJ_NAME,PRJ_COMP,PRJ_DEPT,PRJ_PHASE_NAME,CREATE_TIME,S,FOLDER)=res
                if (PrjName,comp,dept,PrjPhaseName,CreateTime,S_NAME,FolderName)!=res and CreateTime >= CREATE_TIME:
                    cur.execute(u"UPDATE CAPOL_PRJACSSERVER SET PRJ_NAME=:1,PRJ_COMP=:2,PRJ_DEPT=:3,PRJ_PHASE_NAME=:4,CREATE_TIME=:5,S_NAME=:6,FOLDER=:7 where PRJ_NO=:8 and PrjPhaseId=:9",[PrjName,comp,dept,PrjPhaseName,CreateTime,S_NAME,FolderName,PrjNo,PrjPhaseId])
            else:
                cur.execute(u"INSERT INTO CAPOL_PRJACSSERVER(PRJ_NO,PRJ_NAME,PRJ_COMP,PRJ_DEPT,PRJ_PHASE_NAME,PrjPhaseId,CREATE_TIME,S_NAME,SAP_PRJNO,FOLDER) VALUES(:1,:2,:3,:4,:5,:6,:7,:8,:9,:10)",[PrjNo,PrjName,comp,dept,PrjPhaseName,PrjPhaseId,CreateTime,S_NAME,PrjNo,FolderName])
            cur.execute(u"UPDATE CAPOL_PRJACSSERVER SET PRJ_NAME='最新同步时间',CREATE_TIME=TRUNC(SYSDATE),S_NAME='0000' where PRJ_NO='0000'")
        self.oracle.commit()

    def run_sync(self,drop=0):
        if drop:
            self.drop_and_create()
        self.get_project()
        self.check_and_insert()
        

class staff_sync(sync):

    def get_staff(self):
        cur=self.cip.cursor()
        cur.execute(u"""SELECT LoginName,JobNo,Username,CompanyName,DeptName,Specialty,Specialty as SubDept FROM [dbo].[v_auditor_staff]""")
        self.staff=cur.fetchall()
        return self.staff

    def drop_and_create(self):
        cur=self.oracle.cursor()
        try:
            cur.execute(u"DROP TABLE CAPOL_CIP_STAFF")
        except cx_Oracle.DatabaseError:
            pass
        cur.execute(u"CREATE TABLE CAPOL_CIP_STAFF(LoginName varchar2(50),JobNo varchar2(16),Username varchar2(50),CompanyName varchar2(50),DeptName varchar2(50),Specialty varchar2(50),SubDept varchar2(50))")
        cur.execute(u"INSERT INTO CAPOL_CIP_STAFF(LoginName,JobNo,Username) VALUES('同步时间','0000',to_char(sysdate, 'yyyy-mm-dd hh24:mi:ss'))")
        self.oracle.commit()

    def check_and_insert(self):
        staff=self.get_staff()
        cur=self.oracle.cursor()
        for (LoginName,JobNo,Username,CompanyName,DeptName,Specialty,SubDept) in staff:
            cur.execute(u"SELECT LoginName,JobNo,Username,CompanyName,DeptName,SubDept FROM CAPOL_CIP_STAFF WHERE JobNo=:1",[JobNo])
            res=cur.fetchone()
            if res:
                if (LoginName,JobNo,Username,CompanyName,DeptName,SubDept)!=res:
                    cur.execute(u"UPDATE CAPOL_CIP_STAFF SET LoginName=:1,Username=:2,CompanyName=:3,DeptName=:4,SubDept=:5 WHERE JobNo=:6",[LoginName,Username,CompanyName,DeptName,SubDept,JobNo])
            else:
                cur.execute(u" INSERT INTO CAPOL_CIP_STAFF(LoginName,JobNo,Username,CompanyName,DeptName,Specialty,SubDept) VALUES(:1,:2,:3,:4,:5,:6,:7)",[LoginName,JobNo,Username,CompanyName,DeptName,Specialty,SubDept])
        cur.execute(u"UPDATE CAPOL_CIP_STAFF SET Username=to_char(sysdate, 'yyyy-mm-dd hh24:mi:ss') where JobNo='0000'")
        self.oracle.commit()

    def run_sync(self,drop=0):
        if drop:
            self.drop_and_create()
        self.get_staff()
        self.check_and_insert()

class project_role_sync(sync):

    def __init__(self, *arg, **kw):
        sync.__init__(self, *arg, **kw)
        self.prj_role=[]

    def run_sync(self,drop=1):
        if drop:
            self.drop_and_create()
        cur=self.cip.cursor()
        cur.execute(u"""SELECT JobNo,Username,fd_login_name,GCNO,GCName,PhaseName,Gwname FROM v_auditor_project_role""")
        res=cur.fetchone()
        cur2=self.oracle.cursor()
        while res:
            cur2.execute(u"INSERT INTO CAPOL_SAP_PRJ_ROLE(JobNo,Username,fd_login_name,GCNO,GCName,PhaseName,Gwname) VALUES(:1,:2,:3,:4,:5,:6,:7)" ,res)
            res=cur.fetchone()
        cur2.execute(u"UPDATE CAPOL_SAP_PRJ_ROLE SET fd_login_name=to_char(sysdate, 'yyyy-mm-dd hh24:mi:ss') WHERE JobNo='0000'")
        self.oracle.commit()

    def run_sync1(self,drop=1):
        if drop:
            self.drop_and_create()
        cur=self.cip.cursor()
        cur.execute(u"""SELECT JobNo,Username,fd_login_name,GCNO,GCName,PhaseName,Gwname FROM v_auditor_project_role""")
        cur2=self.oracle.cursor()
        for res in cur.fetchall():
            cur2.execute(u"INSERT INTO CAPOL_SAP_PRJ_ROLE(JobNo,Username,fd_login_name,GCNO,GCName,PhaseName,Gwname) VALUES(:1,:2,:3,:4,:5,:6,:7)" ,res)
        cur2.execute(u"UPDATE CAPOL_SAP_PRJ_ROLE SET fd_login_name=to_char(sysdate, 'yyyy-mm-dd hh24:mi:ss') WHERE JobNo='0000'")
        self.oracle.commit()

    def drop_and_create(self):
        cur=self.oracle.cursor()
        try:
            cur.execute(u"DROP TABLE CAPOL_SAP_PRJ_ROLE")
        except cx_Oracle.DatabaseError:
            pass
        cur.execute(u"""CREATE TABLE CAPOL_SAP_PRJ_ROLE(JobNo varchar2(16),Username varchar2(50),fd_login_name varchar2(50),
            GCNO varchar2(16),GCName varchar2(200),PhaseName varchar2(16),Gwname varchar2(50))""")
        cur.execute(u"INSERT INTO CAPOL_SAP_PRJ_ROLE(JobNo,Username,fd_login_name) VALUES('0000','同步时间',to_char(sysdate, 'yyyy-mm-dd hh24:mi:ss'))")
        self.oracle.commit()



if __name__=="__main__":
    run()
