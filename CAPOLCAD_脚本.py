# -*- coding: utf-8 -*-
import pymssql,cx_Oracle
import os


os.environ['NLS_LANG'] = 'SIMPLIFIED CHINESE_CHINA.UTF8'
method='capol'
#method='test'

def run():
    for i in GetTasks('符润红'):
        print(i)

def GetTasks(fd_name):
    with sync(method=method) as s:
        lst=[]
        cur=s.cip.cursor()
        cur.execute("SELECT fd_id from cip.dbo.sys_org_element where fd_name=%s " ,(fd_name))
        (fd_id,)=cur.fetchone()
        cur.execute("SELECT f_project_no,f_project_name,f_head_personId from cip_tl.dbo.t_workspace_project")
        for (f_project_no,f_project_name,f_head_personId) in cur.fetchall():
            if fd_id in f_head_personId.split(";"):
                lst.append((f_project_no,f_project_name,f_head_personId))
    return lst


def FlushBlockInsertRecord():
    with sync(method=method) as s:
        cur=s.oracle.cursor()
        cur.execute("SELECT ID,BLOCKNAME from CAPOL_BLOCKINSERTRECORD")
        res=cur.fetchall()
        for (ID,BLOCKNAME) in res:
            #print(filepath)
            #print(filepath.split("\\")[-1])
            if BLOCKNAME[-4:]!='.dwg':
                new_name=BLOCKNAME+'.dwg'
                cur.execute("UPDATE CAPOL_BLOCKINSERTRECORD set BLOCKNAME=:1 WHERE ID=:2",(new_name,ID))
            #filename=filepath.split("\\")[-1]
           # print(filename)
         #   cur.execute("UPDATE CAPOL_STANDARDCADFILEPATH set FILEPATH=:1 WHERE ID=:2",(filename,ID))
        s.oracle.commit()

def FlushStandartCADFile():
    with sync(method=method) as s:
        cur=s.oracle.cursor()
        cur.execute("SELECT ID,FILEPATH from CAPOL_STANDARDCADFILEPATH ")
        res=cur.fetchall()
        for (ID,filepath) in res:
            #print(filepath)
            #print(filepath.split("\\")[-1])
            filename=filepath.split("\\")[-1]
            print(filename)
            cur.execute("UPDATE CAPOL_STANDARDCADFILEPATH set FILEPATH=:1 WHERE ID=:2",(filename,ID))
        s.oracle.commit()

def fill_capol_projectopinions_2018():
    with sync(method=method) as s:
        cur=s.oracle.cursor()
        for a in ['S','C']:
            for b in ['J','G','S','N','D']:
                cur.execute(u"""SELECT opid,aa,bb,tt from (SELECT 意见编号 as opid,
                    case when 工程阶段='施工图' then 'S' when 工程阶段='初步设计' then 'C' else '0' end as aa,
                    case when 专业='建筑' then 'J' WHEN 专业='结构' THEN 'G'when 专业='给排水' then 'S' WHEN 专业='弱电' or 专业='电气' THEN 'D'when 专业='暖通' then 'N' else 'O' end as bb,
                    校审时间 as tt from capol_projectopinions_2018) t where aa=:1 and bb=:2 order by tt asc""",[a,b])
                res=cur.fetchall()
                i=1
                for (opid,aa,bb,tt) in res:
                    ii='000000'+str(i)
                    n_opid=b+a+'YJ'+ii[-4:]
                    cur.execute(u"UPDATE capol_projectopinions_2018 SET 意见编码=:1,及时校审=1,及时修改=1,是否有效=1 where 意见编号=:2",(n_opid,opid))
                    #print(opid,aa,bb,tt,n_opid)
                    i+=1
        s.oracle.commit()
        cur.execute(u"select PRJ_NO,capol_companyserverpath.ID from capol_prjacsserver left join capol_localcompanyserverpath on capol_companyserverpath.company=capol_prjacsserver.S_NAME")
        for (PRJ_NO,serverid) in cur.fetchall():
            cur.execute(u"UPDATE capol_projectopinions_2018 SET 项目所在公司=:1 where 工程编号=:2",(serverid,PRJ_NO))
        s.oracle.commit()
        cur.execute(u"SELECT USERNAME,JOBNO FROM capol_cip_staff")
        for (USERNAME,JOBNO) in cur.fetchall():
            cur.execute(u"UPDATE capol_projectopinions_2018 SET 意见责任人工号=:1 where 意见责任人=:2",(JOBNO,USERNAME))
            cur.execute(u"UPDATE capol_projectopinions_2018 SET 校审人工号=:1 where 校审人=:2",(JOBNO,USERNAME))
        s.oracle.commit()

class sync():
    def __init__(self,method='test'): 
        self.cip=pymssql.connect(host='10.1.1.117', user="cip_ro", password="hygj!@#4567",database="cip")
        if method=='test':
            oracle='steven/qfc23834358@10.1.42.131:1521/xe'
        elif method=='capol':
            oracle='ccd/hygj1234@10.1.1.213:1521/orcl'
        else:
            oracle=''
        if oracle:
            self.oracle=cx_Oracle.connect(oracle)
        self.acs_server={}
        self.acs_server_lst=[('深圳','10.1.246.1'),('广州','10.2.1.114'),('长沙','10.3.1.3'),('上海','10.6.1.240')]
        for (S_NAME,host) in self.acs_server_lst:
            self.acs_server[S_NAME]=pymssql.connect(host=host, user="readonly", password="capol!@#456",database="CAPOL_Project")

    def __enter__(self):
        return self

    def __exit__(self, type, value, trace):
        self.a=1
        self.cip.close()
        self.oracle.close()
        for S_NAME in self.acs_server:
            self.acs_server[S_NAME].close()

if __name__=="__main__":
    run()
