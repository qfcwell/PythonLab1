# -*- coding: utf-8 -*-
import pymssql,time

def run():
    with WH_Conn() as app:
        app.run()

class WH_Conn():
    def __init__(self): 
        self.cip=pymssql.connect(host='10.1.1.117', user="cip_ro", password="hygj!@#4567",database="cip")
        self.test2=pymssql.connect(host='10.1.42.103', user="sa", password="qfc23834358Q",database="test2")

    def __enter__(self):
        return self

    def __exit__(self, type, value, trace):
        self.close()

    def close(self):
        self.cip.close()
        self.test2.close()

    def GetProjectStages(self):
        cur=self.cip.cursor()
        cur.execute("""SELECT prj.fd_project_no,prj.fd_project_name,whp.fd_ppoject_stage,comp.fd_name as comp,dept.fd_name as dept 
            from wt_workhours_project whp join wt_project_info prj on whp.fd_project_id=prj.fd_id
            left join sys_org_element comp on comp.fd_id=prj.doc_company_id
            left join sys_org_element dept on dept.fd_id=prj.doc_dept_id 
            where  prj.doc_create_time>'2010-01-01' group by prj.fd_project_no,prj.fd_project_name,whp.fd_ppoject_stage,comp.fd_name,dept.fd_name """)
        return cur.fetchall()

    def drop_and_create(self):
        cur=self.test2.cursor()
        cur.execute("""DROP TABLE wh_Table""")
        cur.execute("""CREATE TABLE [dbo].[wh_Table]([id] [smallint] IDENTITY(1,1) NOT NULL,[prjno] [nvarchar](50) NULL,[prjname] [nvarchar](50) NULL,[stage] [nvarchar](50) NULL,[comp] [nvarchar](50) NULL,[dept] [nvarchar](50) NULL,[wh] [float],[rate] [float],[isf] [int] NULL,[y] [int] NULL,[m] [int] NULL) ON [PRIMARY]""")
        self.test2.commit()
        return 0

    def run(self):
        self.drop_and_create()
        for (PrjNo,PrjName,Stage,comp,dept) in self.GetProjectStages():
            print(PrjNo+PrjName+'-'+Stage)
            wh=WH_prj(self.cip,self.test2,PrjNo,PrjName,Stage,comp,dept)

class WH_prj():
    def __init__(self,cip,test2,PrjNo,PrjName,Stage,comp,dept): 
        self.cip=cip
        self.test2=test2
        self.C_Y,self.C_M=time.localtime()[0],time.localtime()[1]
        self.PrjNo=PrjNo
        self.PrjName=PrjName
        self.Stage=Stage
        self.comp=comp
        self.dept=dept
        self.WH_Total=self.GetWH_Total() 
        if self.WH_Total:
            self.GetFinish()
            self.GetRate()
            self.save()
        else:
            print(PrjNo+Stage+'无工时')

    def GetPrjInfo(self):
        cur=self.cip.cursor()
        cur.execute("SELECT fd_project_name,comp.fd_name as comp,dept.fd_name as dept, FROM wt_project_info prj left join sys_org_element comp on comp.fd_id=prj.doc_company_id left join sys_org_element dept on dept.fd_id=prj.doc_dept_id WHERE prj.fd_project_no=%s",(self.PrjNo))
        res=cur.fetchone()
        if res:
            (self.PrjName,self.comp,self.dept)=res
            return 1
        else:
            (self.PrjName,self.comp,self.dept)=('','','')
            return 0

    def GetWH_Total(self):
        cur=self.cip.cursor()
        cur.execute("SELECT sum(fd_std_workhours) from wt_workhours_project whp JOIN wt_project_info prj on whp.fd_project_id=prj.fd_id where prj.fd_project_no=%s and whp.fd_ppoject_stage=%s ",(self.PrjNo,self.Stage))
        res=cur.fetchone()
        if res:
            return res[0]
        else:
            return 0

    def GetFinish(self,FinishPoint=0.9):
        cur=self.cip.cursor()
        cur.execute("""SELECT y,m,sum(fd_std_workhours) from(
                select fd_workhours_date,fd_user_no,fd_ppoject_stage,fd_workhours_num,fd_std_workhours,
                prj.fd_project_name,prj.fd_project_no,prj.doc_create_time,comp.fd_name as comp,dept.fd_name as dept,
                YEAR(fd_workhours_date) as y,MONTH(fd_workhours_date) as m
                from wt_workhours_project whp 
                join wt_project_info prj on whp.fd_project_id=prj.fd_id
                left join sys_org_element comp on comp.fd_id=prj.doc_company_id
                left join sys_org_element dept on dept.fd_id=prj.doc_dept_id
                where fd_project_no=%s AND fd_ppoject_stage=%s) as t1
                group by y,m
                order by y,m asc""",(self.PrjNo,self.Stage))
        t1=self.WH_Total*FinishPoint
        t2=0.0
        res=cur.fetchall()
        for (y,m,wh) in res:
            if wh:
                t2+=wh
            if t2>t1:
                break
        if y*12+m<self.C_Y*12+self.C_M-3:
            self.y,self.m=y,m
            self.PrjIsFinished=1
        else:
            self.y,self.m=0,0
            self.PrjIsFinished=0

    def GetRate(self):
        if self.Stage in ['投标','规划','方案','方案配合','初设','初设配合','施工图','施工服务','BIM','课题']:
            cur=self.test2.cursor()
            sql="SELECT [%(stage)s] from wh_Table2 where [工程编号]='%(prjno)s'" % {'stage':self.Stage,'prjno':self.PrjNo}
            #print(sql)
            cur.execute(sql)
            res=cur.fetchone()
            if res:
                self.rate=res[0]
            else:
                self.rate=0
        else:
            self.rate=0

    def save(self):
        cur=self.test2.cursor()
        cur.execute("""INSERT INTO wh_Table([prjno],[prjname],[stage],[comp],[dept],[wh],[rate],[isf],[y],[m]) VALUES(%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)""",(self.PrjNo,self.PrjName,self.Stage,self.comp,self.dept,self.WH_Total,self.rate,self.PrjIsFinished,self.y,self.m))
        self.test2.commit()

if __name__=="__main__":
    run()
