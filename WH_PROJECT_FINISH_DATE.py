# -*- coding: utf-8 -*-
import pymssql,time

def run():
    app=WH_Conn()
    app.run()
    '''for Stage in app.stages:
        wh=app.GetWH(PrjNo='GC150016',Stage=Stage)
        print(Stage+str(wh))
        app.GetFinish(PrjNo='GC150016',Stage=Stage,FinishPoint=0.9)'''


class WH_Conn():
    def __init__(self): 
        self.cip=pymssql.connect(host='10.1.1.117', user="cip_ro", password="hygj!@#4567",database="cip")
        self.test2=pymssql.connect(host='10.1.42.103', user="sa", password="qfc23834358Q",database="test2")
        self.C_Y,self.C_M=time.localtime()[0],time.localtime()[1]
        self.stages=self.GetStages()
        self.ProjectStages=self.GetProjectStages()

    def __enter__(self):
        return self

    def __exit__(self, type, value, trace):
        self.close()

    def close(self):
        self.cip.close()
        self.test2.close()

    def GetProjects(self):
        cur=self.cip.cursor()
        cur.execute("SELECT fd_id,fd_project_no,fd_project_name from wt_project_info where doc_create_time>'2015-01-01'")
        self.Projects=cur.fetchall()
        return self.Projects

    def GetStages(self):
        stages=[]
        cur=self.cip.cursor()
        cur.execute("SELECT fd_ppoject_stage from wt_workhours_project group by fd_ppoject_stage")
        for (stage,) in cur.fetchall():
            stages.append(stage)
        return stages

    def GetProjectStages(self):
        cur=self.cip.cursor()
        cur.execute("""SELECT prj.fd_project_no,prj.fd_project_name,whp.fd_ppoject_stage,comp.fd_name as comp,dept.fd_name as dept 
            from wt_workhours_project whp join wt_project_info prj on whp.fd_project_id=prj.fd_id
            left join sys_org_element comp on comp.fd_id=prj.doc_company_id
            left join sys_org_element dept on dept.fd_id=prj.doc_dept_id 
            where  prj.doc_create_time>'2015-01-01' group by prj.fd_project_no,prj.fd_project_name,whp.fd_ppoject_stage,comp.fd_name,dept.fd_name """)
        return cur.fetchall()

    def GetWH(self,PrjNo,Stage):
        cur=self.cip.cursor()
        cur.execute("SELECT sum(fd_std_workhours) from wt_workhours_project whp JOIN wt_project_info prj on whp.fd_project_id=prj.fd_id where prj.fd_project_no=%s and whp.fd_ppoject_stage=%s ",(PrjNo,Stage))
        res=cur.fetchone()
        if res:
            return res[0]
        else:
            return 0

    def GetFinish(self,PrjNo,Stage,FinishPoint=0.9):
        cur=self.cip.cursor()
        WH_Total=self.GetWH(PrjNo,Stage)
        if WH_Total:
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
                    order by y,m asc""",(PrjNo,Stage))
            t1=WH_Total*FinishPoint
            t2=0.0
            res=cur.fetchall()
            for (y,m,wh) in res:
                if wh:
                    t2+=wh
                if t2>t1:
                    break
            if y*12+m<self.C_Y*12+self.C_M-3:
                return (y,m)
            else:
                return '未完成'

        else:
            print(PrjNo+Stage+'无工时')
            return '无工时'

    def save(self,PrjNo,PrjName,Stage,comp,dept,PrjIsFinished,y,m):
        cur=self.test2.cursor()
        cur.execute("""INSERT INTO wh_Table([prjno],[prjname],[stage],[comp],[dept],[isf],[y],[m]) VALUES(%s,%s,%s,%s,%s,%s,%s,%s)""",(PrjNo,PrjName,Stage,comp,dept,PrjIsFinished,y,m))
        self.test2.commit()

    def drop_and_create(self):
        cur=self.test2.cursor()
        cur.execute("""DROP TABLE wh_Table""")
        cur.execute("""CREATE TABLE [dbo].[wh_Table]([id] [smallint] IDENTITY(1,1) NOT NULL,[prjno] [nvarchar](50) NULL,[prjname] [nvarchar](50) NULL,[stage] [nvarchar](50) NULL,[comp] [nvarchar](50) NULL,[dept] [nvarchar](50) NULL,[isf] [int] NULL,[y] [int] NULL,[m] [int] NULL) ON [PRIMARY]""")
        self.test2.commit()
        return 0

    def run(self):
        self.drop_and_create()
        for (PrjNo,PrjName,Stage,comp,dept) in self.ProjectStages:
            res=self.GetFinish(PrjNo,Stage)
            print(PrjNo+PrjName+'-'+Stage)
            if res!='无工时':
                if res=='未完成':
                    PrjIsFinished,y,m=0,0,0
                else:
                    (y,m)=res
                    PrjIsFinished=1
                self.save(PrjNo,PrjName,Stage,comp,dept,PrjIsFinished,y,m)

if __name__=="__main__":
    run()
